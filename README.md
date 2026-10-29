# IntelliRequest Community

A public-source, **local-only** Word document update tool. Analysts import their own DOCX, enter field values or explicit text replacements, preview the amended document, and save a *new local copy*. No institutional templates or case data are distributed by this repository.

## Install in PyCharm (Windows)
1. Clone https://github.com/jctc-ng/intellirequest-community.git.
2. Configure Python **3.11** in PyCharm.
3. In the project terminal: `python -m pip install -r requirements.txt`.
4. Run **server.py**, not the project folder.
5. Visit **http://127.0.0.1:8765** in Chrome or Edge.

For a page-faithful PDF preview, install and license Microsoft Word (desktop) and the Windows Python dependency pywin32. Without Word, a basic content/table preview remains available. No LibreOffice needed.

## Workflow
1. Select your DOCX from your own computer.
2. Enter the requested fields; placeholders such as `{{EXCHANGE}}`, `{{ENTITY}}`, `{{REFERENCE}}`, `{{OFFICER}}`, `{{SUBJECTS}}` and `{{DATE_LONG}}` can be filled.
3. For documents without placeholders, add explicit Find/Replace rules.
4. Generate and review the modified document preview.
5. Choose **Save updated document locally**. You decide where on your device to save the changed DOCX. The original stays untouched.

The tool does not automatically infer hidden content or guarantee a clean document in any legal sense. All changes require review and authorization; legal statements and signer details must be confirmed by the operator.

## Security
The backend is bound to localhost only. No network requests, analytics, or external conversion services are used to process documents. Do not put institutional templates, case data, addresses or credentials in GitHub. Public source code does not grant access to templates held on analysts' computers. However, software cannot stop a person who possesses a document from copying it using other applications.

## Update in PyCharm
Stop server, run `git pull origin main`, run `python -m pip install -r requirements.txt`, restart `server.py`, refresh the browser.

## Tests
Run `python -m unittest discover -v` using the project's Python interpreter.
