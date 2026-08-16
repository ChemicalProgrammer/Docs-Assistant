# Docs Assistant v0.7.7

Google Docs bound/add-on Apps Script starter project.

## Included
- Sidebar with Formatting and Gemini tabs.
- Applies existing Google Docs named paragraph styles; it does not redefine them.
- Paragraph spacing and Keep with Next.
- Bullet / number / letter / Roman list presets.
- UI checkbox reserved for continue/restart numbering.
- Page break insertion.
- Gemini assistant with preview before applying.
- Insert at cursor or replace current selection.
- Quick translate, summarize, expand, technical and concise actions.
- Per-user Gemini API key in UserProperties.
- Settings dialog.

## Important MVP limitations
1. Section breaks are exposed in the UI but need implementation through the advanced Google Docs API.
2. Exact continuation/restart behavior for complex numbered lists needs the Google Docs API; DocumentApp alone is limited.
3. Gemini replacement currently targets textual selections. Rich multi-paragraph/table formatting preservation should be implemented as a dedicated document-range engine.
4. "Intelligent formatting" (Gemini classifying headings/tables/normal text and then applying styles) is intentionally left for the next module rather than letting Gemini directly mutate the document.

## Install
Create or open an Apps Script project attached to a Google Doc, copy these files, save, reload the document, and use:
Add-ons / Extensions → Docs Assistant → Open Docs Assistant.

For a published Workspace Add-on, manifest/deployment configuration will need to be adapted to the chosen deployment model.

## API key
Open Settings from the add-on menu or sidebar. The key is stored with:
PropertiesService.getUserProperties()

It is not written into the document or source code.


## v0.2.0 UI
- Workspace-inspired card UI and gradient Gemini accent.
- Global execution lock: only one task can run at a time.
- All controls disabled during server/Gemini operations.
- Animated busy overlay and active-button shimmer.
- Success/error micro-interactions.
- Reduced-motion accessibility support.


## v0.3.0 UI
- Removed global loading overlay.
- Only active button animates during execution.
- All other controls are disabled until completion.
- Colored functional button families.
- Colored white-text tabs for Formatting and Gemini.
- Inline spinner, shimmer, Done/Error feedback in active button.


## v0.4.0
- Smart Insert / Replace: pasted text is semantically classified by Gemini and applied with native Docs headings, normal paragraphs, lists, and tables.
- Selection automatically means replace; otherwise cursor means insert. No instruction is required.
- Settings dialog redesigned to match sidebar styling, colors and button execution animations.


## v0.4.1
- Fixed multi-paragraph selection detection. Normal text and Heading 1–6 now apply to every paragraph touched by the selection.


## v0.5.0
- Added Format selected table.
- Black 1 pt borders, no background, 0.49 in minimum row height, 0.028 in padding, middle vertical alignment.
- First row: Arial 9, bold, 1.5 line spacing, centered.
- Body: Arial 9, regular, 1.5 line spacing, left aligned, vertically centered.
- Column widths are left automatic by not forcing a width.
- Properties not exposed reliably by DocumentApp are not simulated.


## v0.5.1
- List formatting now first applies the document's current Normal text named style.
- Existing list items are reused instead of recreated.
- Bullet/number/letter/Roman list items use:
  - Left indent: 0.06 in
  - Hanging indent: 0.25 in
- Hanging indent is implemented as first-line position 0.06 in and wrapped-line start 0.31 in.


## v0.5.2
- Heading/Normal buttons now read the document's current named-style attributes with `Body.getHeadingAttributes()`.
- The selected paragraphs receive those attributes explicitly after applying the named style, including the named style's current indentation and paragraph spacing.
- Rich-text attributes from the current named style are also reapplied to prevent prior direct character formatting from overriding the style visually.
- Heading 1–6 text is converted to sentence case automatically: first alphabetic character uppercase, remaining text lowercase.
- Normal text does not change capitalization.


## v0.5.3
- Heading indentation is now applied explicitly after the current named style:
  - Heading 1: Left -0.12 in, Right 0 in, Special indent None.
  - Heading 2: Left 0 in, Right 0 in, Special indent None.
  - Heading 3: Left 0.19 in, Right 0 in, Special indent None.
- Heading 1–3 are explicitly left aligned.
- Numeric heading prefixes typed into the text are normalized to exactly one regular space after the number (for example `1.2.3   Title` -> `1.2.3 Title`).
- Headings 4–6 keep their current named-style indentation until explicit values are defined.


## v0.5.4
- Normal text and Heading 1–6 no longer require selecting the paragraph text.
- With no selection, the paragraph containing the cursor is treated as the complete target line.
- With a multi-paragraph selection, styles still apply to every selected paragraph.
- Heading sentence case, current named-style attributes, and explicit heading indents continue to apply to the whole target paragraph.


## v0.5.5
- Added Figure/Table caption button.
- Input format: `Figure X. Description` or `Table X. Description` (existing numeric values also accepted).
- Caption formatting:
  - Normal text named style as base.
  - Arial 9.
  - Centered paragraph.
  - Only `Figure N.` / `Table N.` is bold.
- Figure and Table numbering are independent integer sequences.
- The correct number is calculated by scanning all prior captions of the same type in document order, rather than relying on the immediately previous caption.


## v0.6.0
- Fixed caption parsing: `Table 5.1. Overview` now becomes `Table N. Overview`; the complete old dotted identifier is removed.
- Bullet, numeric, letter/inciso and Roman list buttons now share one formatting engine:
  - current Normal text style first
  - Left indent 0.06 in
  - Hanging indent 0.25 in
- Added `Format complete selection`.
  - Gemini is called once to classify the selected paragraphs.
  - Formatting is then applied locally with the existing deterministic functions.
  - Detects Normal text, Heading 1–6, bullets, numeric lists, letter incisos, Roman lists and Figure/Table captions.
  - Actual table-cell content is deliberately skipped in the full-format pass; use `Format selected table` for table objects.
  - Figure/Table caption lines remain supported and are numbered using a single document scan for better performance.


## v0.6.1
- Table formatting now applies Custom Spacing to all text inside all cells:
  - Before: 0 pt
  - After: 0 pt
- Applies to header and body cells, including paragraph and list-item content.


## v0.6.2
- Table cells now use single line spacing (1.0) in both header and body cells.
- Custom spacing remains 0 pt before and 0 pt after.


## v0.6.3
- All table-cell paragraphs and list items now use:
  - Left indent: 0.05 in
  - Right indent: 0 in
  - Special indent: None
- Existing table rules remain:
  - Single line spacing
  - 0 pt before / 0 pt after


## v0.6.4
- Figure/Table caption detection is now fully case-insensitive (`FIGURE`, `Figure`, `figure`, `TABLE`, etc.).
- A caption can begin with only the keyword; an existing X/number is optional.
- Decimal/chapter-style identifiers are consumed as one complete old number:
  - `Table 5.1. Overview` -> `Table N. Overview`
  - `Table 5.1 Overview` -> `Table N. Overview`
  - `Figure 5.1.3. Diagram` -> `Figure N. Diagram`
- The parser also recovers from malformed prior output such as `Table 54. 1. Overview` by consuming `54. 1` as the old identifier.


## v0.6.5
- Split the caption action into separate `Figure` and `Table` buttons.
- A caption keyword is no longer required in the source line.
  - `Overview of the process` + Figure -> `Figure N. Overview of the process`
  - `5.1. Overview` + Table -> `Table N. Overview`
- Existing Figure/Table prefixes and old dotted numbers are removed before applying the new consecutive number.
- Button text is smaller throughout the sidebar.
- Compact buttons now use a minimal `…` loading label instead of long labels such as `Applying…`.


## v0.7.0
- Added a separate `Note` button.
  - Accepts a line with or without Note/Notes/Nota/Notas.
  - Output is `Note. Description`.
  - Base Normal text style, Arial 9, centered.
  - Only `Note.` is bold.
  - No counter.
- Rebuilt `Format complete selection` as a hybrid formatter:
  - deterministic local detection first;
  - Gemini only classifies ambiguous paragraphs;
  - ambiguous paragraphs are processed in batches for long selections;
  - a malformed Gemini JSON batch no longer aborts the whole operation;
  - omitted Gemini items safely fall back to Normal text.
- Full Format now detects existing native headings/lists, decimal section headings, bullets, letter incisos, Roman lists, Figure/Table captions and Notes locally.
- Manual list prefixes such as `a)`, `i)`, `1.` and `•` are removed before converting text into native Google Docs lists, avoiding duplicated markers.
- Actual table-cell content remains excluded from Full Format and continues to use the dedicated table formatter.


## v0.7.1
- Note formatting now removes existing Note/Notes/Nota/Notas prefixes before applying the normalized `Note.` prefix.
- Balanced outer parentheses are removed when the whole note is wrapped, e.g. `(Note: Abcdefg...)` -> `Note. Abcdefg...`.
- If no Note/Nota prefix exists, the current line is preserved as the description and `Note.` is added.
- Added spacing below the Figure/Table/Note button row so its description aligns visually with the other sections.


## v0.7.2
- Note parser rebuilt:
  - `(Note: Abcdefg...)` -> `Note. Abcdefg...`
  - `(Nota: Abcdefg...)` -> `Note. Abcdefg...`
  - existing Note/Notes/Nota/Notas plus wrappers/punctuation are removed;
  - if no Note/Nota marker exists, the original line is preserved and only `Note.` is prepended.
- Lists:
  - letter incisos are rendered explicitly as `a)`, `b)`, `c)` to guarantee the required suffix;
  - native numeric lists keep the `1.`, `2.`, `3.` convention;
  - all list types keep Left 0.06 in / Hanging 0.25 in.
- Continue previous numbering:
  - native number/Roman/bullet lists search upward for the nearest compatible ListItem and reuse its list ID;
  - letter incisos search upward for the previous letter and continue from it;
  - Restart creates a new native list for numeric/Roman/bullet lists and starts letter incisos at `a)`.
- Safety guards:
  - Table formatting only runs when the cursor/selection is entirely inside one actual table.
  - Figure/Table caption and Note formatting are blocked inside table cells.
  - Full Smart Format continues to skip table-cell content.


## v0.7.3
- Fixed neighboring numbering/heading corruption when formatting bullets or native lists.
- Root cause: a newly inserted Google Docs ListItem can temporarily inherit the listId of a nearby multilevel heading/outline list. Changing its glyph before detaching it can modify that shared list scheme.
- Native list workflow now:
  1. finds the intended compatible previous list when Continue is enabled, or creates an isolated temporary listId when restarting;
  2. assigns the safe listId first;
  3. only then applies Bullet/Number/Roman glyph and indentation.
- Continue numbering now also checks the expected Add-on indentation (Left 0.06 in / Hanging 0.25 in), preventing numbered headings from being mistaken for the preceding list.
- Full Smart Format's native-list helper uses the same isolation safeguard.


## v0.7.4
- Replaced native-list mutation with a safe manual-list engine.
- The Add-on no longer calls `setGlyphType()` or `setListId()` when applying Bullet, Number, Letter or Roman formatting.
- This prevents list formatting from changing the shared multilevel-list definition used by nearby numbered headings.
- Visual list prefixes are explicit:
  - Bullet: `•`
  - Number: `1.`, `2.`, `3.`
  - Letter: `a)`, `b)`, `c)`
  - Roman: `i.`, `ii.`, `iii.`
- Continue numbering scans upward for the previous matching visible prefix.
- Older native lists created by previous versions are supported read-only for continuation when they match the Add-on indentation.
- Selected native ListItems are converted only into plain paragraphs; neighboring list definitions are never modified.


## v0.7.5
- Fixed the first manually formatted list paragraph inheriting bold from the line above.
- Manual list formatting now reapplies the document's current Normal text style after replacing the paragraph text.
- If Normal text is not itself bold, any inherited whole-paragraph bold is explicitly cleared.
- The same fix is used by Full Smart Format.


## v0.7.7
- Built from the stable v0.7.5 list engine; v0.7.6 list/indent changes were rolled back.
- List buttons now work with only the cursor in the current line.
- If the target is already a compatible native Google Docs list item:
  - it is NOT converted to a paragraph;
  - listId is untouched;
  - glyph type is untouched;
  - existing native-list indentation is untouched;
  - automatic sequence below it is preserved.
- Plain paragraphs still use the stable v0.7.5 manual-list fallback and its existing Left 0.06 in / Hanging 0.25 in behavior.
- Full Smart Format uses the same native-list preservation rule.
