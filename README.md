# Tanooki DOCX Converter

Restyles short-utterance transcript `.docx` files (DOC_01) into the DOC_02 layout. It keeps every line and every word exactly as they are and only changes how each line looks. Everything runs in the browser, so files are never uploaded anywhere.

**In:** one paragraph per utterance

```
01:00:08:00    Transcribed    Contributor 2 - Right, are we rolling?
01:00:12:20    Transcribed    Contributor 2 - Okay.
```

**Out:** the same paragraphs in the DOC_02 layout. The speaker name is bold, the timecode is grey (`#777777`), and the text is plain Helvetica.

> **Contributor 2** [01:00:08] Right, are we rolling?  
> **Contributor 2** [01:00:12] Okay.

## Rules

- One line in gives one paragraph out. Wording, order and speakers stay the same.
- By default the timecode is trimmed to `[HH:MM:SS]`, as in DOC_02. Tick **Keep frames** to show `[HH:MM:SS:FF]` instead.
- The title lines above the first timecode are kept. Blank paragraphs inside the transcript are dropped.
- Any other paragraph is kept as it is.
- Styles, fonts and page setup come through from the source file unchanged. Only `word/document.xml` is rewritten.
- The output file is named `<original name>_FORMATTED.docx`.

Tested on a real 1,238-line transcript: every line came through word for word, with the right formatting on every paragraph.

## Publish on GitHub Pages

1. Create a new repository, for example `tanooki-docx-converter`.
2. Upload `index.html`, `convert.js` and `README.md` to the root of the repo.
3. Go to **Settings → Pages**, set **Source** to *Deploy from a branch*, pick `main` and `/ (root)`, then save.
4. After about a minute the site is live at `https://<your-username>.github.io/tanooki-docx-converter/`.

`index.html` already contains the converter inline. `convert.js` is the same engine as a standalone file, which also runs in Node:

```js
const JSZip = require('jszip');
const { convertDocx } = require('./convert.js');
convertDocx(JSZip, fs.readFileSync('in.docx'), { nodebuffer: true /*, frames: true */ })
  .then(r => fs.writeFileSync('out.docx', r.file));
```

The engine also has a `merge: true` option that joins consecutive lines from the same speaker, which reproduces the reference DOC_02 exactly. The page doesn't offer it.

The only external dependency is JSZip 3.10.1, loaded from cdnjs.
