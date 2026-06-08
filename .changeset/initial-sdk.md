---
'@schnsrw/casual-sheets': minor
---

Initial release: `@schnsrw/casual-sheets` SDK shipping the signing pipeline
(drawn / typed / uploaded signature surfaces, sequential / concurrent modes)
and the iframe postMessage protocol (`EmbedTransport`, `casual.*` envelope
types). Wire shapes are byte-identical to `@schnsrw/docx-js-editor` — only
the `app` discriminator (`'sheet'` vs `'docs'`) and signature anchor shape
(`{ kind: 'sheet', sheet, cell }` vs `{ kind: 'doc', paraId }`) differ. The
Univer-Sheets React wrapper (`CasualSheets` component) is planned for a
follow-up release.
