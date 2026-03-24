# n8n-nodes-docx2md

An n8n community node for converting between **DOCX**, **Markdown**, and **HTML** formats — in both directions.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/sustainable-use-license/) workflow automation platform.

[Installation](#installation)
[Operations](#operations)
[Compatibility](#compatibility)
[Usage](#usage)
[Resources](#resources)
[Version history](#version-history)

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

> **Note:** This node uses external npm dependencies (`mammoth`, `marked`, `html-to-docx`) and therefore runs on **self-hosted n8n only**. It is not compatible with n8n Cloud.

## Operations

| Operation           | Input                                                   | Output                                                  |
| ------------------- | ------------------------------------------------------- | ------------------------------------------------------- |
| **DOCX → Markdown** | Binary DOCX file                                        | Markdown text (JSON field) + optional binary `.md` file |
| **DOCX → HTML**     | Binary DOCX file                                        | HTML text (JSON field) + optional binary `.html` file   |
| **Markdown → DOCX** | Markdown text (typed, JSON field, or binary `.md` file) | Binary DOCX file                                        |
| **HTML → DOCX**     | HTML text (typed, JSON field, or binary `.html` file)   | Binary DOCX file                                        |

### DOCX → Markdown

Converts a DOCX binary file to Markdown text using [mammoth](https://github.com/mwilliamson/mammoth.js).

| Parameter                    | Description                                                                                                            |
| ---------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| Input Binary Field           | Name of the binary field holding the DOCX file (default: `data`)                                                       |
| Output JSON Field            | JSON property name to store the Markdown string (default: `markdown`)                                                  |
| Extract Images               | When enabled, extracts embedded images; the output field becomes `{ markdown, images }` with `![](image_1)` references |
| Output Binary .md File       | Also attach the Markdown as a binary `.md` file                                                                        |
| Output Markdown Binary Field | Binary field name for the output `.md` file (default: `markdown`)                                                      |

### DOCX → HTML

Converts a DOCX binary file to HTML using [mammoth](https://github.com/mwilliamson/mammoth.js).

| Parameter                | Description                                                                                                 |
| ------------------------ | ----------------------------------------------------------------------------------------------------------- |
| Input Binary Field       | Name of the binary field holding the DOCX file (default: `data`)                                            |
| Output JSON Field        | JSON property name to store the HTML string (default: `html`)                                               |
| Extract Images           | When enabled, extracts embedded images; the output field becomes `{ html, images }` with `image_1` src keys |
| Output Binary .html File | Also attach the HTML as a binary `.html` file                                                               |
| Output HTML Binary Field | Binary field name for the output `.html` file (default: `html`)                                             |

### Markdown → DOCX

Converts Markdown text to a DOCX file using [marked](https://marked.js.org/) + [html-to-docx](https://github.com/privateOmega/html-to-docx).

| Parameter                                           | Description                                                                                 |
| --------------------------------------------------- | ------------------------------------------------------------------------------------------- |
| Markdown Source                                     | Where to read Markdown from: `Enter Markdown` (editor), `JSON Text Field`, or `Binary File` |
| Markdown / Markdown Field Name / Input Binary Field | Depends on the chosen source                                                                |
| Output DOCX Binary Field                            | Binary field name for the generated DOCX file (default: `data`)                             |
| Output File Name                                    | File name for the DOCX (default: `output.docx`)                                             |

### HTML → DOCX

Converts HTML text to a DOCX file using [html-to-docx](https://github.com/privateOmega/html-to-docx).

| Parameter                                   | Description                                                                         |
| ------------------------------------------- | ----------------------------------------------------------------------------------- |
| HTML Source                                 | Where to read HTML from: `Enter HTML` (editor), `JSON Text Field`, or `Binary File` |
| HTML / HTML Field Name / Input Binary Field | Depends on the chosen source                                                        |
| Output DOCX Binary Field                    | Binary field name for the generated DOCX file (default: `data`)                     |
| Output File Name                            | File name for the DOCX (default: `output.docx`)                                     |

## Compatibility

- Requires **self-hosted n8n** (not compatible with n8n Cloud due to external dependencies)
- Tested with n8n v1.x

## Usage

**Convert a DOCX attachment to Markdown:**

1. Add a trigger that receives a DOCX file (e.g. Email trigger or HTTP webhook)
2. Add the **DOCX ↔ Markdown ↔ HTML** node
3. Set Operation to `DOCX → Markdown`
4. Set Input Binary Field to the binary field name holding the DOCX
5. The output item will contain a `markdown` JSON field with the converted text

**Generate a DOCX from Markdown:**

1. Provide Markdown text via any upstream node
2. Add the **DOCX ↔ Markdown ↔ HTML** node
3. Set Operation to `Markdown → DOCX`
4. Set Markdown Source to `JSON Text Field` and enter the field name
5. The output item will contain a binary DOCX file in the configured output field

## Resources

- [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
- [mammoth.js – DOCX to HTML/Markdown](https://github.com/mwilliamson/mammoth.js)
- [marked – Markdown parser](https://marked.js.org/)
- [html-to-docx – HTML to DOCX converter](https://github.com/privateOmega/html-to-docx)

## Version history

### 0.1.0

Initial release. Supports four operations:

- DOCX → Markdown
- DOCX → HTML
- Markdown → DOCX
- HTML → DOCX
