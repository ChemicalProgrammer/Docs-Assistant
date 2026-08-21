# Docs Assistant v0.9.5

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


## v0.7.8
- Existing native Google Docs list items now receive the required local indentation while preserving automatic numbering:
  - Left: 0.06 in
  - Hanging: 0.25 in
  - Right: 0 in
- The ListItem itself is preserved: listId and glyph type are not changed.
- Cursor-only formatting remains supported.


## v0.7.9
- Fixed Table/Figure caption counters using document child indexes instead of JavaScript object identity.
- Root cause: comparing Apps Script document element wrappers with `===` is not reliable; the target caption was not always detected in document order.
- Table captions now derive their number from the nearest ACTUAL Google Docs table element and that table's ordinal in the document.
- Table caption may be above or below the table.
- If no actual table exists, Table numbering falls back to prior Table captions.
- Figure numbering continues to count prior Figure captions.
- Full Smart Format caption numbering uses the same corrected counter.


## v0.8.0
- Figure captions now derive their consecutive number from actual visual objects, similar to Table captions.
- A standalone Figure block is a body-level paragraph/list item containing:
  - InlineImage
  - InlineDrawing
  - PositionedImage anchored to the paragraph
- Multiple visual objects in one paragraph count as one composite Figure.
- Graphics inside tables are not counted as standalone Figures.
- The Figure caption is associated with the nearest visual block; on equal distance, the visual block above the caption is preferred.
- If no actual visual object is found, numbering falls back to prior Figure captions.
- Table numbering remains based on actual TABLE elements.


## v0.8.1
- Bullet formatting now creates REAL Google Docs `ListItem` bullets instead of inserting the `•` text character.
- Existing native bullet items are preserved; their listId/glyph are not changed.
- Plain paragraphs are converted to native bullets with an isolation workflow that prevents inheriting or mutating nearby numbered-heading list definitions.
- Required bullet indentation remains Left 0.06 in / Hanging 0.25 in / Right 0.
- Continue previous numbering/list for bullets joins the nearest safe native bullet list above.
- Full Smart Format also uses real native bullets.


## v0.8.2
- Added configurable caption numbering anchors for Tables and Figures.
- New sidebar controls under CAPTIONS & NOTES:
  - Start at [N]
  - Set here
  - Clear
  - Renumber captions
- `Set here` finds the nearest actual Table/Figure object to the cursor and stores a Named Range anchor in the document.
- Start values are stored in Document Properties.
- Tables/Figures before an anchor are excluded from the custom sequence.
- `Clear` removes the anchor and returns that counter to document-start counting at 1.
- `Renumber captions` updates all existing body-level Figure/Table captions; captions before a configured anchor are left unchanged.
- Table counting uses actual TABLE elements.
- Figure counting uses actual standalone image/drawing blocks.
- This solves cover/layout tables being counted before the real report content.


## v0.8.3
- Fixed Caption Numbering anchor setup error ("no table near cursor").
- `Set here` now anchors the CURRENT caption paragraph instead of trying to locate the physical Table/Figure during setup.
- At numbering time, the Add-on resolves that caption anchor to the nearest actual Table/Figure:
  - Table: prefers the table after the caption.
  - Figure: prefers the figure before the caption.
- This is more robust for captions separated from objects by blank paragraphs, page breaks, spacing, or other body elements.
- Workflow: place cursor on the first real caption line -> Start at N -> Set here.


## v0.8.4
- Fixed false "The cursor must be in the main document body" errors when setting caption anchors.
- Root cause: the code compared Apps Script element wrapper objects with JavaScript identity (`===`) while walking from the cursor paragraph to the Body. The same document element can be represented by separate wrappers.
- Body membership is now detected by `DocumentApp.ElementType.BODY_SECTION` and validated with `Body.getChildIndex()`.
- Caption NamedRanges are now created/read/removed through the active `DocumentTab` when available (`tab.newRange()`, `tab.addNamedRange()`, `tab.getNamedRanges()`), with legacy Document fallback.
- The user workflow remains: cursor on first caption -> Start at -> Set here.


## v0.8.5 — Performance + equations
- Performance:
  - Named-style attributes are loaded once per formatting operation and reused across selected paragraphs.
  - Removed redundant named-style lookups and one redundant heading assignment.
  - Text-style attributes are filtered once per operation.
  - Sidebar controls unlock as soon as the Apps Script call completes; the previous fixed 700 ms post-success lock was removed.
  - Success feedback remains visible briefly (180 ms) without blocking the next action.
  - Simple formatting actions no longer trigger caption-numbering reads after every click.
- Equations:
  - Added `∑ Format equation row`.
  - Uses a borderless 3-cell table with symmetric side columns so the equation remains centered on the page.
  - Deep-copies the original paragraph so embedded Google Docs Equation elements are preserved.
  - Adds a dotted leader and bold/italic `Equation N` in the right cell.
  - Automatically increments from prior equation rows.
  - Removes borders and cell padding for a cleaner technical layout.


## v0.8.6 — Complete headings + table-cell bullets
- Heading 1–6:
  - the complete heading is forced to Bold and 10 pt;
  - this includes manually typed numbering and the title text;
  - paragraph-level Bold/10 pt is also applied so native numbered-heading glyphs inherit the same formatting;
  - all other named-style properties remain based on the document's current Heading style.
- Bullets:
  - the existing Bullet button now detects whether each target is inside a TableCell;
  - normal body bullets remain Left 0.06 in / Hanging 0.25 in;
  - bullets inside table cells use Left 0.00 in / Hanging 0.20 in;
  - native bullet listId/glyph are preserved;
  - Continue previous bullet detection uses the correct indentation for body vs table-cell context;
  - Full Smart Format uses the same table-aware bullet indentation.


## v0.8.7 — Fast paragraph formatting architecture
- Rebuilt Normal/Heading formatting around two explicit routes:
  - `applyNamedStyleToCurrentParagraph(styleName)` for one cursor paragraph.
  - `applyNamedStyleToSelectedParagraphs(styleName)` for a selected paragraph range.
- Both routes use the same low-level `formatSingleParagraph_()` function.
- Multi-paragraph formatting now includes blank paragraphs between the first and last selected paragraph.
- `applyNamedStyle(styleName)` automatically dispatches to single-paragraph vs range mode.
- Major reduction in DocumentApp mutations:
  - one `setHeading()`;
  - one paragraph `setAttributes()`;
  - one text `setAttributes()`;
  - optional `setText()` only when a heading actually needs sentence-case/number-spacing normalization.
- H1/H2/H3 indentation is folded into the single paragraph-attributes call rather than four separate setters.
- Heading sentence case + numeric-prefix spacing are calculated in memory and written at most once.
- Spacing/Keep-with-next now use the same cursor-or-inclusive-selection paragraph engine.
- Sidebar status now reports Apps Script server execution time for style buttons, helping distinguish code execution from Apps Script/network startup latency.


## v0.8.8 — Uniform left alignment
- Normal text and Heading 1–6 now use the same paragraph geometry:
  - Left indent: 0 in
  - Right indent: 0 in
  - Special indent: None
  - Alignment: Left
- Removed the previous H1/H2/H3 custom left indents (-0.12 / 0 / 0.19 in).
- The alignment override is merged into the existing paragraph `setAttributes()` call, so it adds no extra DocumentApp mutations.


## v0.8.9 — Reliable Normal / Heading formatter
- Rebuilt the Normal + Heading 1–6 formatter after v0.8.8 regressions.
- Removed `Paragraph.setAttributes()` from the named-style path; mixing paragraph and character attributes there was the main reliability risk.
- Single-paragraph and multi-paragraph paths are explicit:
  - `applyNamedStyleToCurrentParagraph(styleName)`
  - `applyNamedStyleToSelectedParagraphs(styleName)`
- Both use the same `formatSingleParagraph_()` implementation.
- Named style is applied with `setHeading()`.
- Paragraph geometry is applied explicitly:
  - Left 0
  - Right 0
  - First line 0
  - Left alignment
- Normal text reapplies the document Normal style's character attributes so inherited bold/direct formatting is cleared.
- Heading 1–6 reapplies named-style character attributes, then forces the complete line to Bold + 10 pt.
- Table cells are protected from Normal/Heading buttons.


## v0.9.0 — Named styles simplified to native Google Docs behavior
- Normal text and Heading 1–6 now do exactly one formatting operation per paragraph: `paragraph.setHeading(...)`.
- No `getHeadingAttributes()` calls.
- No forced font size, bold, capitalization, indentation, alignment, or text attributes.
- The current document Named Style is therefore the source of truth, exactly as configured in Google Docs.
- Restored `applyNamedStyleToParagraph_()` as the single low-level formatter used throughout the Add-on. This also fixes callers in Lists, Notes/Captions, Gemini Smart Format, etc. that still depended on that helper.
- The style buttons call one server method (`applyNamedStyle`) and the server decides whether the current target is a cursor paragraph or a selection. This avoids relying on stale sidebar selection state.
- Selection formatting includes blank paragraphs between the first and last selected paragraph.
- Timing diagnostics now report:
  - server execution time;
  - total client-observed time.
  If server time is small but total time is large, the remaining delay is Apps Script / `google.script.run` startup or transport latency, not the formatting function itself.


## v0.9.1 — Indentation controls
- Added an Indentation section to the Formatting sidebar inspired by the native Google Docs indentation dialog.
- Controls:
  - Left (inches)
  - Right (inches)
  - Special indent: None / First line / Hanging
  - By (shown only for First line/Hanging)
  - Load current
  - Apply
- The function works with either the cursor in one paragraph or a multi-paragraph selection.
- It changes only indentation; it does not alter Named Style, font, size, bold, capitalization, or alignment.
- Geometry:
  - None: first line = left, wrapped lines = left
  - First line: first line = left + by, wrapped lines = left
  - Hanging: first line = left, wrapped lines = left + by


## v0.9.2 — Fast Apply + optional Reset Overrides
- Named-style buttons now have two explicit modes:
  - Fast Apply (default): exactly one native `setHeading()` formatting mutation per paragraph; no style-property reads.
  - Reset direct formatting (optional checkbox): reads the current Named Style once, then reapplies its effective paragraph/text properties to remove common direct overrides.
- Cursor is checked BEFORE selection. This prevents an old/stale Docs selection from turning a one-paragraph click into a large expensive range operation.
- Reset mode merges Normal-text attributes with the selected Heading attributes so inherited Heading properties are included.
- Granular performance diagnostics now report:
  - target lookup time
  - named-style lookup time
  - formatting-apply time
  - total Apps Script server time
  - total client-observed time
- Reset is intentionally OFF by default because the native Apply path is the fastest possible sidebar implementation.


## v0.9.3 — Simple A/B/C formatting architecture
- A) `applyStyleToParagraph_(paragraph, styleName)`
  - lightweight reset: Normal text -> requested style;
  - no style-property reads;
  - no manual font/size/indent reconstruction.
- B) `getSegments_(selection, comparison)`
  - classifies selected content as Tables, Blank lines, H1-H6, Lists/Bullets/Numbers/Letters/Romans, Normal paragraphs, Figures, Equations, Figure captions, Table captions, Notes, etc.;
  - with no selection, only the cursor element is considered;
  - filters include `NORMAL_PARAGRAPH`, `HEADING`, `STYLEABLE_TEXT`, `LIST`, `TABLE`, `FIGURE`, etc.
- C) `applyStyleToSelection_(selection, styleName)`
  - gets `STYLEABLE_TEXT` segments using B;
  - calls A for each returned paragraph.
- Sidebar uses one public call: `applyStyleToCurrentContext(styleName)`.
- Existing callers remain compatible through `applyNamedStyleToParagraph_()` alias.


## v0.9.4 — Active document body helper restored
- Fixed the runtime error `getActiveBody_ is not defined`.
- Restored the two central document-tab helpers:
  - `getActiveDocumentTab_()`
  - `getActiveBody_()`
- `getActiveBody_()` uses the current Google Docs tab through
  `getActiveTab().asDocumentTab().getBody()`, with a legacy `doc.getBody()`
  fallback.
- No changes were made to the A/B/C style architecture.


## v0.9.5 — Numbered heading classification fix
- Fixed `No styleable paragraph was found at the cursor/selection`.
- Google Docs numbered headings are commonly `ListItem` objects that also have a Heading style.
- Segment classification now checks H1–H6 BEFORE classifying an element as Bullet/Number/Letter/Roman.
- Added a direct cursor fallback: if segmentation returns nothing, the owning Paragraph/ListItem under the cursor is formatted directly.
- No additional style inspection or reconstruction was added.
