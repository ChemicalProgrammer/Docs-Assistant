# Docs Assistant v0.2.0

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
