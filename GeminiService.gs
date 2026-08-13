const GEMINI_MODEL = 'gemini-2.5-flash';

function processWithGemini(instruction, rawText) {
  const context = getDocumentContext();
  const source = String(rawText || '').trim() || context.selectedText;

  const systemInstruction = [
    'You are a writing assistant embedded in Google Docs.',
    'Follow the requested transformation only.',
    'Do not add commentary, preambles, markdown fences, or explanations.',
    'Return only the document-ready result.',
    'Preserve factual meaning unless the instruction explicitly asks otherwise.'
  ].join('\n');

  const prompt = [
    systemInstruction,
    '',
    'TASK:',
    String(instruction || 'Improve the writing while preserving meaning.'),
    '',
    'CONTENT:',
    source
  ].join('\n');

  return callGemini_(prompt);
}

function quickGeminiAction(action, rawText) {
  const actions = {
    TRANSLATE_EN: 'Translate to English. Preserve meaning, terminology, numbers, units, and paragraph structure.',
    TRANSLATE_ES: 'Translate to Spanish. Preserve meaning, terminology, numbers, units, and paragraph structure.',
    SUMMARIZE_50: 'Summarize to approximately 50% of the original length while preserving the important technical information.',
    SUMMARIZE_75: 'Summarize to approximately 75% of the original length while preserving the important technical information.',
    EXPAND_125: 'Expand to approximately 125% of the original length. Add clarity but do not invent facts.',
    EXPAND_150: 'Expand to approximately 150% of the original length. Add clarity but do not invent facts.',
    TECHNICAL: 'Rewrite in clear professional technical language. Do not invent facts.',
    CONCISE: 'Make the text more concise while preserving all important information.'
  };
  if (!actions[action]) throw new Error('Unknown Gemini action.');
  return processWithGemini(actions[action], rawText);
}

function callGemini_(prompt) {
  const key = getGeminiApiKey_();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
              encodeURIComponent(GEMINI_MODEL) + ':generateContent?key=' +
              encodeURIComponent(key);

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2 }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const data = JSON.parse(response.getContentText() || '{}');
  if (code < 200 || code >= 300) {
    throw new Error(data?.error?.message || ('Gemini request failed: HTTP ' + code));
  }

  const text = data?.candidates?.[0]?.content?.parts
    ?.map(p => p.text || '').join('').trim();

  if (!text) throw new Error('Gemini returned no text.');
  return text;
}

function insertGeminiResult(text, mode) {
  const doc = DocumentApp.getActiveDocument();
  text = String(text || '');
  if (!text) throw new Error('There is no Gemini result to insert.');

  if (mode === 'REPLACE') {
    const selection = doc.getSelection();
    if (!selection) throw new Error('No selection is available to replace.');

    const ranges = selection.getRangeElements();
    if (!ranges.length) throw new Error('Selection is empty.');

    // Safe MVP: replace textual selection only, leaving content outside the
    // selection untouched. Complex multi-element formatting comes next.
    ranges.forEach((re, index) => {
      const el = re.getElement();
      if (!el.editAsText) return;
      const t = el.asText();
      if (re.isPartial()) {
        const start = re.getStartOffset();
        const end = re.getEndOffsetInclusive();
        t.deleteText(start, end);
        if (index === 0) t.insertText(start, text);
      } else if (index === 0) {
        t.setText(text);
      } else {
        t.setText('');
      }
    });
    return true;
  }

  const cursor = doc.getCursor();
  if (!cursor) throw new Error('Place the cursor in the document.');
  cursor.insertText(text);
  return true;
}


/**
 * Smart Apply:
 * Gemini classifies pasted content into semantic blocks, then Apps Script
 * applies native Google Docs structures/styles at the current selection/cursor.
 * Gemini never receives permission to edit the document itself.
 */
function smartApplyGeminiContent(rawText) {
  rawText = String(rawText || '').trim();
  if (!rawText) throw new Error('Paste content first.');

  const prompt = [
    'Analyze the following content for insertion into a Google Docs document.',
    'Do NOT rewrite, summarize, translate, correct, or invent content.',
    'Only classify and structure the supplied content.',
    'Return ONLY valid JSON. No markdown fences.',
    'Schema:',
    '{"blocks":[{"type":"normal|heading1|heading2|heading3|heading4|heading5|heading6|bullet|number|table","text":"...","rows":[["..."]]}]}',
    'Rules:',
    '- Preserve the original wording and values.',
    '- Use heading types only when the content clearly functions as a heading.',
    '- normal = ordinary paragraphs.',
    '- bullet/number = list items.',
    '- table = genuinely tabular content; put cells in rows and omit text.',
    '- Preserve block order.',
    '',
    'CONTENT:',
    rawText
  ].join('\n');

  let response = callGemini_(prompt).trim();
  response = response.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  let data;
  try { data = JSON.parse(response); }
  catch (e) { throw new Error('Gemini could not return a valid document structure. Try again.'); }
  if (!data.blocks || !Array.isArray(data.blocks)) throw new Error('Invalid Gemini document structure.');

  applyStructuredBlocks_(data.blocks);
  return { ok:true, blocks:data.blocks.length };
}

function applyStructuredBlocks_(blocks) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  const cursor = doc.getCursor();

  // If text is selected, delete only the textual selection first.
  // Then use the first selected paragraph as the insertion anchor.
  let anchor = null;
  if (selection) {
    const ranges = selection.getRangeElements();
    if (ranges.length) {
      let first = ranges[0].getElement();
      while (first && first.getType() !== DocumentApp.ElementType.PARAGRAPH &&
             first.getType() !== DocumentApp.ElementType.LIST_ITEM) first = first.getParent();
      anchor = first;
      for (let i=ranges.length-1;i>=0;i--) {
        const re=ranges[i], el=re.getElement();
        if (!el.editAsText) continue;
        const t=el.asText();
        if (re.isPartial()) t.deleteText(re.getStartOffset(), re.getEndOffsetInclusive());
        else t.setText('');
      }
    }
  }

  // Insert through a temporary marker at cursor when no selection.
  if (!anchor) {
    if (!cursor) throw new Error('Place the cursor where the content should be inserted.');
    const marker = cursor.insertText('\uE000');
    anchor = marker.getParent();
  }

  const parent = anchor.getParent();
  let index = parent.getChildIndex(anchor);

  // Remove marker if present, while preserving surrounding text.
  try {
    const at=anchor.asParagraph().editAsText();
    const txt=at.getText();
    const mi=txt.indexOf('\uE000');
    if(mi>=0) at.deleteText(mi,mi);
  } catch(e){}

  const headingMap = {
    heading1:DocumentApp.ParagraphHeading.HEADING1,
    heading2:DocumentApp.ParagraphHeading.HEADING2,
    heading3:DocumentApp.ParagraphHeading.HEADING3,
    heading4:DocumentApp.ParagraphHeading.HEADING4,
    heading5:DocumentApp.ParagraphHeading.HEADING5,
    heading6:DocumentApp.ParagraphHeading.HEADING6
  };

  blocks.forEach(block => {
    const type=String(block.type||'normal').toLowerCase();
    if(type==='table' && Array.isArray(block.rows) && block.rows.length){
      const table=parent.insertTable(++index, block.rows.map(r=>r.map(c=>String(c??''))));
      return;
    }
    const text=String(block.text??'');
    if(type==='bullet' || type==='number'){
      const li=parent.insertListItem(++index,text);
      li.setGlyphType(type==='bullet'?DocumentApp.GlyphType.BULLET:DocumentApp.GlyphType.NUMBER);
      return;
    }
    const p=parent.insertParagraph(++index,text);
    p.setHeading(headingMap[type] || DocumentApp.ParagraphHeading.NORMAL);
  });

  // Remove an empty anchor created by replacement when safe.
  try {
    if(anchor.getText && anchor.getText()==='' && parent.getNumChildren()>1) anchor.removeFromParent();
  } catch(e){}
}


function classifyFormattingPlanWithGemini_(items) {
  const prompt = [
    'Classify paragraphs from a technical Google Docs document.',
    'Do NOT rewrite any text. Return only JSON.',
    'Allowed types:',
    'normal, heading1, heading2, heading3, heading4, heading5, heading6, bullet, number, letter, roman, figure_caption, table_caption.',
    '',
    'Rules:',
    '- Preserve every id exactly.',
    '- If fixedType is not empty, return that exact type.',
    '- heading1 = main numbered section/title, often 1. Title.',
    '- heading2 = subsection, often 1.1 Title.',
    '- heading3 = deeper subsection, often 1.1.1 Title.',
    '- heading4/5/6 = progressively deeper headings when clearly applicable.',
    '- bullet = unordered bullet item.',
    '- number = ordered numeric list item that is NOT a section heading.',
    '- letter = inciso/list item such as a), b), c), or equivalent.',
    '- roman = roman numeral list item such as i., ii., iii. or i), ii).',
    '- normal = ordinary prose.',
    '- Use surrounding paragraphs to distinguish numbered headings from numbered lists.',
    '',
    'Return exactly this shape:',
    '{"items":[{"id":"p0","type":"normal"}]}',
    '',
    'INPUT:',
    JSON.stringify(items)
  ].join('\n');

  let response = callGemini_(prompt).trim();
  response = response.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');

  let data;
  try {
    data = JSON.parse(response);
  } catch (e) {
    throw new Error('Gemini could not classify the selected document structure.');
  }

  if (!data.items || !Array.isArray(data.items)) {
    throw new Error('Gemini returned an invalid formatting plan.');
  }

  const allowed = {
    normal:true,
    heading1:true, heading2:true, heading3:true,
    heading4:true, heading5:true, heading6:true,
    bullet:true, number:true, letter:true, roman:true,
    figure_caption:true, table_caption:true
  };

  return data.items
    .filter(x => x && typeof x.id === 'string' && allowed[String(x.type || '').toLowerCase()])
    .map(x => ({id:x.id, type:String(x.type).toLowerCase()}));
}
