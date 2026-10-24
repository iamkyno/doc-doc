# DocDoc

A simple, friendly document editor for Windows. Word-style ribbon, the usual keyboard shortcuts, and it saves real `.docx` files.

## Install

1. Go to the **Releases** page of this repo.
2. Download `DocDoc-Setup-x.y.z.exe` and run it.
3. Windows may show "Windows protected your PC" (the app isn't code-signed). Click **More info → Run anyway**.

## What it does

- **Home:** cut/copy/paste, format painter, fonts and sizes, bold/italic/underline/strikethrough, sub/superscript, change case, text colour, highlight, bullets, numbering, indent, alignment, line spacing, paragraph shading, styles (Normal, Headings, Title, Quote), find & replace
- **Insert:** page break, tables, pictures (also paste or drag in; pictures pasted from web pages are downloaded and embedded), 18 shapes, text boxes, links, horizontal line, date, symbols
- **Table Layout tab** (appears when the cursor is in a table, also on right-click): insert rows/columns, delete row/column/table, merge and split cells, cell shading, text position in cells, all/outside/no borders, table alignment and width
- **Undo / Redo** for every change (typing, formatting, tables, boxes, shape formatting), with buttons at the top left
- **Find & Replace** highlights every match and shows "2 of 5"
- **Open by drag and drop:** drop a document anywhere on the window
- **Shape Format tab** (appears when a text box, shape or picture is selected): fill colour, no fill, transparency, border colour/weight/style (solid, dashed, dotted, double), rounded corners, shadow, rotation, inner margin, text position, exact size in cm, wrapping, front/back, delete
- **Exam paper boxes** (Insert → Shapes): mark box, mark out of (side by side, stacked, or circle), mark allocation (5), mark circle, total box, tick boxes (click to tick), answer lines, dotted lines, answer box, examiner's-use table, candidate details with number boxes, instructions box. Mark boxes stay beside their question as you edit
- **Paragraph borders:** bottom, top, box around paragraphs
- **Layout:** margins, orientation, page size, columns, page colour
- **Page view:** real separate pages on screen; print and PDF break pages at the same places
- **Headers, footers and page numbers:** left / centre / right text, `{page}`, `{pages}`, `{date}`, `{title}`, different first page; saved to Word and read from Word files
- **Autosave & recovery:** a backup is kept while there are unsaved changes and offered after a crash
- **Recent files:** File tab → Recent and File menu → Open Recent
- **View:** zoom (or Ctrl + mouse wheel), page view on/off, spell check
- **Files:** open `.docx`, `.docm`, `.dotx`, `.html`, `.md`, `.txt` directly. Open `.doc`, `.rtf`, `.odt` when Microsoft Word or the free LibreOffice is installed (DocDoc uses it to convert). Save as `.docx`, `.html`, `.txt`; export PDF; print
- **Existing Word documents** keep their fonts, sizes, colours, highlights, alignment, spacing, indents, headings, bullet/numbered lists, tables (borders and shading), pictures, page size and margins
- **Fonts:** every font installed on your PC, plus 30 built-in free fonts (Roboto, Open Sans, Lato, Montserrat, Poppins, Inter, Merriweather, Playfair Display, Lora, EB Garamond, Dancing Script, Fira Code and more). Carlito and Caladea stand in for Calibri and Cambria on PCs without Office. Type in the font box to search

Press **F1** in the app to see every keyboard shortcut.

## Updates

Installed copies of DocDoc check GitHub for a newer version when they start (and every 4 hours).
When one is found, users get an **Update available** popup with the release notes. The update
downloads in the background, then **Restart & update** installs it (or it installs the next time
DocDoc is closed). Users can also check any time with **Help → Check for Updates** or **File → Check for Updates**.

### Publishing an update

1. Make your changes.
2. Raise `"version"` in `package.json` (for example `1.5.0` → `1.5.1`). Users only get an update when this number goes up.
3. Write what's new in `RELEASE_NOTES.md` (this text appears in the popup).
4. Push to `master`. GitHub Actions builds the installer and publishes the release `v1.5.1` with `latest.yml`, which is what the app checks.

Pushes without a version change still build the installer, but don't create a release, so nobody is offered a half-finished version.

## Develop

```
npm install
npm start          # run the app
npm run dist       # build the Windows installer (run on Windows)
npm run vendor     # refresh bundled fonts and the .docx renderer in src/
```

Every push to `master` builds the installer on GitHub Actions and attaches it to a release.
