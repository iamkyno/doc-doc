# DocDoc

A simple, friendly document editor for Windows. Word-style ribbon, the usual keyboard shortcuts, and it saves real `.docx` files.

## Install

1. Go to the **Releases** page of this repo.
2. Download `DocDoc-Setup-x.y.z.exe` and run it.
3. Windows may show "Windows protected your PC" (the app isn't code-signed). Click **More info → Run anyway**.

## What it does

- **Home:** cut/copy/paste, format painter, fonts and sizes, bold/italic/underline/strikethrough, sub/superscript, change case, text colour, highlight, bullets, numbering, indent, alignment, line spacing, paragraph shading, styles (Normal, Headings, Title, Quote), find & replace
- **Insert:** page break, tables, pictures (also paste or drag in), 18 shapes, text boxes, links, horizontal line, date, symbols
- **Shape Format tab** (appears when a text box, shape or picture is selected): fill colour, no fill, transparency, border colour/weight/style (solid, dashed, dotted, double), rounded corners, shadow, rotation, inner margin, text position, exact size in cm, wrapping, front/back, delete
- **Exam paper boxes** (Insert → Shapes): mark box, mark out of (__ / 5), mark allocation (5), mark circle, total box, tick boxes (click to tick), answer lines, dotted lines, answer box, examiner's-use table, candidate details with number boxes, instructions box. Mark boxes stay beside their question as you edit
- **Paragraph borders:** bottom, top, box around paragraphs
- **Layout:** margins, orientation, page size, columns, page colour
- **View:** zoom (or Ctrl + mouse wheel), spell check
- **Files:** open `.docx`, `.docm`, `.dotx`, `.html`, `.md`, `.txt` directly. Open `.doc`, `.rtf`, `.odt` when Microsoft Word or the free LibreOffice is installed (DocDoc uses it to convert). Save as `.docx`, `.html`, `.txt`; export PDF; print
- **Existing Word documents** keep their fonts, sizes, colours, highlights, alignment, spacing, indents, headings, bullet/numbered lists, tables (borders and shading), pictures, page size and margins
- **Fonts:** every font installed on your PC, plus 30 built-in free fonts (Roboto, Open Sans, Lato, Montserrat, Poppins, Inter, Merriweather, Playfair Display, Lora, EB Garamond, Dancing Script, Fira Code and more). Carlito and Caladea stand in for Calibri and Cambria on PCs without Office. Type in the font box to search

Press **F1** in the app to see every keyboard shortcut.

## Develop

```
npm install
npm start          # run the app
npm run dist       # build the Windows installer (run on Windows)
npm run vendor     # refresh bundled fonts and the .docx renderer in src/
```

Every push to `master` builds the installer on GitHub Actions and attaches it to a release.
