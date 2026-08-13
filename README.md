# Docs Assistant v0.5.2

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
