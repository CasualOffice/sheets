---
"@casualoffice/sheets": minor
---

Chrome (`chrome="full"`): add a menu bar, color pickers, and a navigable name box.

- **Menu bar** (Edit / Insert / Format / Data) above the toolbar — dropdown menus
  dispatching Univer commands (undo/redo, insert row/col, bold/italic/underline,
  sort asc/desc, toggle filter). No logo/title — the host frames the editor.
- **Text & fill color pickers** in the toolbar — swatch popovers (set text color,
  fill color, or reset).
- **Name box** in the formula bar — shows the active cell's A1 reference and jumps
  to a typed cell/range (`B5`, `A1:C3`) on Enter.

All design-system styled, dark-mode aware, driven through `CasualSheetsAPI`.
