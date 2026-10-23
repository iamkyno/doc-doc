# DocDoc

A simple, friendly document editor for Windows. Word-style ribbon, the usual keyboard shortcuts, and it saves real `.docx` files.

## Install

1. Go to the **Releases** page of this repo.
2. Download `DocDoc-Setup-x.y.z.exe` and run it.
3. Windows may show "Windows protected your PC" (the app isn't code-signed). Click **More info → Run anyway**.

## What it does

- **Home:** cut/copy/paste, format painter, fonts and sizes, bold/italic/underline/strikethrough, sub/superscript, change case, text colour, highlight, bullets, numbering, indent, alignment, line spacing, paragraph shading, styles (Normal, Headings, Title, Quote), find & replace
- **Insert:** page break, tables, pictures (also paste or drag in), 18 shapes, text boxes, links, horizontal line, date, symbols
- **Layout:** margins, orientation, page size, columns, page colour, picture/shape wrapping, bring to front/back, fill and outline colours
- **View:** zoom (or Ctrl + mouse wheel), spell check
- **Files:** open `.docx`, `.html`, `.txt`; save as `.docx`, `.html`, `.txt`; export PDF; print

Press **F1** in the app to see every keyboard shortcut.

## Develop

```
npm install
npm start          # run the app
npm run dist       # build the Windows installer (run on Windows)
```

Every push to `master` builds the installer on GitHub Actions and attaches it to a release.
