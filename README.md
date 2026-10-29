# IntelliRequest Community — Browser Edition

**Use IntelliRequest directly in your browser on GitHub Pages.** No Python, PyCharm, LibreOffice, desktop Word, or document-processing server is necessary.

## Access
Once GitHub Pages is enabled and the deployment workflow succeeds, open:
**https://jctc-ng.github.io/intellirequest-community/**

### Enable GitHub Pages once
1. Open https://github.com/jctc-ng/intellirequest-community/settings/pages
2. Under **Build and deployment → Source**, choose **GitHub Actions**.
3. Open **Actions → Publish IntelliRequest Community (browser only)** and check the latest workflow run. If needed, use **Run workflow**.
4. The Pages URL above becomes usable after the successful deployment.

The site is a compiled static application. Its JavaScript dependencies (`jszip` and `docx-preview`) are bundled during the GitHub Actions build, not loaded as third-party runtime scripts.

## Analyst workflow
1. Open the website in Chrome or Edge, on an authorized device.
2. Click **Choose DOCX from this device** and select your own local document.
3. Fill any matching placeholders (such as `{{EXCHANGE}}`, `{{ENTITY}}`, `{{INCORPORATION}}`, `{{REFERENCE}}`, `{{DATE_LONG}}`, `{{OFFICER}}`, `{{SUBJECTS}}`) or add exact Find/Replace rules for documents that do not contain placeholders.
4. Click **Generate document preview** to inspect the modified document.
5. Click **Save updated document on my device**, choose a separate backup filename, and grant the browser permission to update the selected original DOCX.
6. Open the final Word document and verify the layout, all updated fields, attached materials, legal statements, signatures and page breaks.

**The website never supplies a template.** An analyst can only operate on a document they have imported from their own device.

## Privacy and security
- The site is a static HTML/JavaScript bundle hosted on GitHub Pages.
- DOCX archives are parsed, modified, previewed and saved inside the analyst's browser. **No upload endpoint exists.**
- Local document updates rely on the browser's File System Access API and require the user's consent. The program first makes a separate backup.
- The publicly available source code does not include institutional templates, user documents or case records.
- Browsers have access to the document while the page is open. Do not use on shared or untrusted computers. Institutional rules regarding public websites and data classification still apply.
- Public source code and local document possession are not rights management: the software cannot prevent a user from copying files they already possess.

## Rendering limitations
- `docx-preview` is an approximation of Microsoft Word rendering; it may not reproduce every section break, page break, complex table, shape, floating image, field or header with pixel-perfect fidelity.
- Replacements are applied to XML text runs of paragraphs across the Word body, headers, footers and certain notes, preserving other parts of the DOCX.
- Unknown entities hidden in arbitrary text are **not** inferred automatically. Exact placeholders or explicit Find/Replace rules are required.
- The document is not legally certified, signed, approved or transmitted.

## Developing and deploying
Requirements for a developer only: Node.js 22 and npm.

```bash
npm install
npm run build
```

Output in `dist/` is deployed using `.github/workflows/pages.yml` to GitHub Pages whenever changes are pushed to `main`.

## Deprecated local Python version
The old `server.py`, `engine.py` and Python files are maintained as historical prototypes and are **not used by the GitHub Pages edition**. Analysts should use the website, not run Python.

## Troubleshooting the DOCX chooser

- **Choose DOCX from this device** opens a standard browser file input. This works even when direct file editing permissions are unavailable.
- **Open with edit permission** uses the secure File System Access API on compatible Chrome/Edge browsers, allowing the original local document to be updated after a separate backup is saved.
- If direct editing is not permitted, the app offers **Save As** for an updated local copy instead.
- If the script appears stale, refresh with **Ctrl+Shift+R** and check [GitHub Actions](https://github.com/jctc-ng/intellirequest-community/actions/workflows/pages.yml) for the latest successful Pages deployment.
- If the app reports that JavaScript did not load, GitHub Pages may not have deployed the compiled `assets/app.js` from `npm run build`. In GitHub repository Settings → Pages, select **GitHub Actions** as the source.

The browser edition includes the original NFIU crest from the supplied document as a favicon (resized for browser display); no institutional template is bundled.

## Updated form mapping and download workflow

If you import a Word document with **existing text instead of placeholders**, IntelliRequest cannot know which existing company, person, date or jurisdiction is the one you intend to replace. The form therefore provides a **Currently written in document** column beside each input field. Enter the exact original value and the new value in the main form. After import, the app suggests a few detectable old values (e.g., references or dates); verify them carefully before using them.

For any submitted field that is not matched to text in the original document, the default option adds it to a clearly labelled **Information to incorporate / review** section at the end of the generated document. Such information must be moved into its correct place and reviewed before official use. Disable the option if you do not want the review section.

**Download Generated DOCX** creates a new Word file without modifying your imported original. **Update original document (with backup)** is a separate operation requiring explicit file permissions and backup confirmation.

If you still see the older interface, open https://github.com/jctc-ng/intellirequest-community/actions/workflows/pages.yml and confirm the latest Pages workflow succeeded, then force-refresh the website with **Ctrl+Shift+R**.
