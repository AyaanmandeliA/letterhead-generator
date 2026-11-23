# Letterhead generator

A small Node script that produces branded Word (.docx) letters from a JSON brand registry and a JSON content file.

## What it does

You pick a registered letterhead (for example the group letterhead or one brand's own letterhead), point it at a content file, and it writes a finished .docx with the right header, logo treatment, footer, fonts and colours.

## How it works

- `scripts/letterheads.json` is the brand registry. Each entry holds the display name, aliases, logo path and aspect ratio, colours, fonts, masthead line, footer lines and a default signatory.
- `scripts/generate_letterhead.js` reads one registry entry plus a content file and builds the document with the `docx` library:
  - Header: either a centred group wordmark with an optional "issued on behalf of" entity line, or a two-column brand masthead (name on the left, logo on the right). Logos can be shown plain or framed.
  - Footer: either a row of all group brands with small logos, or the brand name with a website and email line.
  - Body: date, optional recipient block, salutation, subject, paragraphs, sign-off and signatory. Short all-caps lines are treated as section headings. Bracketed fill-in fields such as `[Amount]` can be highlighted in yellow.
- `assets/` holds the logo files referenced by the registry.

## Usage

```
cd scripts
npm install
node generate_letterhead.js --config aanka-group --data ../examples/sample-letter.json --out ../output/example-letter.docx
```

`examples/sample-letter.json` is synthetic sample content.

## Stack

Node.js, the `docx` npm package.

## Status

In use internally.

## Removed for privacy

This repository contains only the generator script, its brand config and logo files. Real letters, contracts, signature images, identity documents and internal notes were not included. In the config, the default signatory has been replaced with a placeholder, contact email addresses have been replaced with `contact@example.com`, and internal change notes have been removed.

I set the requirements and check the results. AI tools write most of the code.
