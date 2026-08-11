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
