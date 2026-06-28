#!/usr/bin/env node
/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Generate the home-page template library.
 *
 *   node scripts/build-templates.mjs
 *
 * Writes 10 real .xlsx files into apps/web/public/templates/. Each
 * template is hand-authored to match Excel's expected feel for that
 * use case: real data, sensible number formats, a formula or two,
 * named ranges where they help, and a tab color that matches the
 * card accent on the home page.
 *
 * Re-running overwrites in place — idempotent. The script is the
 * source of truth for what's in the library; the home page reads
 * the matching entries from `apps/web/src/home/templates/registry.ts`.
 */
import ExcelJS from 'exceljs';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('..', import.meta.url)));
const OUT_DIR = resolve(ROOT, 'apps/web/public/templates');
mkdirSync(OUT_DIR, { recursive: true });

// ── Style helpers ──────────────────────────────────────────────────────────

const HEADER_FONT = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
const BORDER = { style: 'thin', color: { argb: 'FFD0D7DE' } };

function applyHeader(row, accentArgb) {
  row.height = 22;
  row.eachCell((cell) => {
    cell.font = HEADER_FONT;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: accentArgb },
    };
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    cell.border = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
  });
}

function setColumns(ws, defs) {
  ws.columns = defs.map((d) => ({
    key: d.key,
    width: d.width,
  }));
}

function zebraDataRows(ws, startRow, endRow, evenArgb) {
  for (let r = startRow; r <= endRow; r++) {
    if ((r - startRow) % 2 !== 0) {
      ws.getRow(r).eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: evenArgb },
        };
      });
    }
  }
}

async function saveWorkbook(wb, filename) {
  const buf = await wb.xlsx.writeBuffer();
  writeFileSync(resolve(OUT_DIR, filename), Buffer.from(buf));
  console.info(`  ✓ ${filename}  (${(buf.byteLength / 1024).toFixed(1)} KB)`);
}

// ── Templates ──────────────────────────────────────────────────────────────

async function personalBudget() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Casual Sheets';
  wb.title = 'Personal Budget';
  const ws = wb.addWorksheet('Budget', {
    properties: { tabColor: { argb: 'FF2E7D32' } },
    views: [{ state: 'frozen', xSplit: 1, ySplit: 1, activeCell: 'B2' }],
  });
  setColumns(ws, [
    { key: 'category', width: 24 },
    { key: 'budget', width: 14 },
    { key: 'jan', width: 12 },
    { key: 'feb', width: 12 },
    { key: 'mar', width: 12 },
    { key: 'avg', width: 12 },
    { key: 'remaining', width: 14 },
  ]);
  ws.addRow(['Category', 'Budget', 'Jan', 'Feb', 'Mar', 'Avg / mo', 'Remaining']);
  applyHeader(ws.getRow(1), 'FF2E7D32');

  const rows = [
    ['Income · Salary', 5400, 5400, 5400, 5400],
    ['Income · Side gig', 600, 480, 620, 560],
    ['Housing · Rent', 1800, 1800, 1800, 1800],
    ['Housing · Utilities', 220, 195, 240, 215],
    ['Groceries', 520, 510, 530, 495],
    ['Dining out', 240, 280, 195, 220],
    ['Transport · Gas', 180, 150, 165, 175],
    ['Transport · Transit', 100, 92, 100, 92],
    ['Health & wellness', 150, 95, 180, 60],
    ['Subscriptions', 120, 118, 118, 122],
    ['Savings', 800, 800, 800, 1000],
    ['Discretionary', 300, 410, 280, 195],
  ];
  rows.forEach((r) => ws.addRow([r[0], r[1], r[2], r[3], r[4], null, null]));
  // Avg + remaining formulas
  for (let r = 2; r <= rows.length + 1; r++) {
    ws.getCell(`F${r}`).value = { formula: `AVERAGE(C${r}:E${r})` };
    ws.getCell(`G${r}`).value = { formula: `B${r}-F${r}` };
  }
  // Number formats
  for (let r = 2; r <= rows.length + 1; r++) {
    ['B', 'C', 'D', 'E', 'F', 'G'].forEach((col) => {
      ws.getCell(`${col}${r}`).numFmt = '"$"#,##0';
    });
  }
  // Style: bold category, zebra rows
  for (let r = 2; r <= rows.length + 1; r++) ws.getCell(`A${r}`).font = { bold: true, color: { argb: 'FF1F2937' } };
  zebraDataRows(ws, 2, rows.length + 1, 'FFE8F5E9');
  // Totals row
  const totalRow = rows.length + 3;
  ws.getCell(`A${totalRow}`).value = 'Net (Income − Expenses)';
  ws.getCell(`A${totalRow}`).font = { bold: true };
  ws.getCell(`B${totalRow}`).value = { formula: `SUM(B2:B3)-SUM(B4:B${rows.length + 1})` };
  ws.getCell(`C${totalRow}`).value = { formula: `SUM(C2:C3)-SUM(C4:C${rows.length + 1})` };
  ws.getCell(`D${totalRow}`).value = { formula: `SUM(D2:D3)-SUM(D4:D${rows.length + 1})` };
  ws.getCell(`E${totalRow}`).value = { formula: `SUM(E2:E3)-SUM(E4:E${rows.length + 1})` };
  ws.getCell(`F${totalRow}`).value = { formula: `AVERAGE(C${totalRow}:E${totalRow})` };
  ['B', 'C', 'D', 'E', 'F'].forEach((col) => {
    ws.getCell(`${col}${totalRow}`).numFmt = '"$"#,##0';
    ws.getCell(`${col}${totalRow}`).font = { bold: true };
  });
  await saveWorkbook(wb, 'personal-budget.xlsx');
}

async function todoList() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'To-Do List';
  const ws = wb.addWorksheet('Tasks', {
    properties: { tabColor: { argb: 'FF1D4ED8' } },
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'task', width: 38 },
    { key: 'status', width: 14 },
    { key: 'priority', width: 12 },
    { key: 'due', width: 14 },
    { key: 'notes', width: 30 },
  ]);
  ws.addRow(['Task', 'Status', 'Priority', 'Due', 'Notes']);
  applyHeader(ws.getRow(1), 'FF1D4ED8');

  const today = new Date();
  const day = (offset) => new Date(today.getTime() + offset * 86400000);
  const rows = [
    ['Draft proposal for Q3 launch', 'In Progress', 'High', day(2), 'Share with team Friday'],
    ['Reply to Sara about contract', 'To Do', 'High', day(1), ''],
    ['Update onboarding deck', 'Done', 'Medium', day(-3), 'Done — slides v3'],
    ['Book dentist appointment', 'To Do', 'Low', day(14), ''],
    ['Renew domain', 'To Do', 'Medium', day(21), 'casualsheets.com'],
    ['Pay credit card', 'Done', 'High', day(-1), ''],
    ['Plan team offsite', 'In Progress', 'Medium', day(28), 'Coordinate with HR'],
    ['Submit expense report', 'To Do', 'Medium', day(4), 'May expenses'],
    ['Write blog post', 'To Do', 'Low', day(10), '"how we ship" series'],
    ['Review pull requests', 'In Progress', 'High', day(0), '#421, #422, #425'],
  ];
  rows.forEach((r) => ws.addRow(r));

  for (let r = 2; r <= rows.length + 1; r++) {
    ws.getCell(`D${r}`).numFmt = 'yyyy-mm-dd';
    // Status color hint
    const status = ws.getCell(`B${r}`).value;
    const fillFor = (s) =>
      s === 'Done' ? 'FFD1F2E1' : s === 'In Progress' ? 'FFFFE9C7' : 'FFEEF2FF';
    ws.getCell(`B${r}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: fillFor(status) },
    };
    ws.getCell(`B${r}`).alignment = { horizontal: 'center' };
    // Priority bold for High
    if (ws.getCell(`C${r}`).value === 'High') ws.getCell(`C${r}`).font = { bold: true, color: { argb: 'FFB91C1C' } };
  }
  zebraDataRows(ws, 2, rows.length + 1, 'FFF8FAFC');
  await saveWorkbook(wb, 'todo-list.xlsx');
}

async function projectTracker() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Project Tracker';
  const ws = wb.addWorksheet('Projects', {
    properties: { tabColor: { argb: 'FF5B21B6' } },
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'project', width: 28 },
    { key: 'owner', width: 16 },
    { key: 'status', width: 14 },
    { key: 'progress', width: 14 },
    { key: 'start', width: 14 },
    { key: 'due', width: 14 },
    { key: 'budget', width: 14 },
    { key: 'spent', width: 14 },
  ]);
  ws.addRow(['Project', 'Owner', 'Status', 'Progress', 'Start', 'Due', 'Budget', 'Spent']);
  applyHeader(ws.getRow(1), 'FF5B21B6');

  const today = new Date();
  const d = (offset) => new Date(today.getTime() + offset * 86400000);
  const rows = [
    ['Onboarding revamp', 'Aria', 'On Track', 0.6, d(-30), d(20), 18000, 9800],
    ['Mobile redesign', 'Liam', 'At Risk', 0.35, d(-45), d(15), 42000, 31000],
    ['Search v2', 'Maya', 'On Track', 0.78, d(-60), d(7), 22000, 14000],
    ['Marketing site', 'Noah', 'Blocked', 0.2, d(-20), d(40), 12000, 3000],
    ['Pricing refresh', 'Aria', 'Done', 1.0, d(-90), d(-2), 8000, 7400],
    ['Billing migration', 'Priya', 'On Track', 0.5, d(-15), d(60), 28000, 11000],
    ['Analytics MVP', 'Sam', 'At Risk', 0.42, d(-25), d(18), 16000, 9200],
    ['Internal LMS', 'Jordan', 'On Track', 0.3, d(-10), d(50), 9500, 2400],
  ];
  rows.forEach((r) => ws.addRow(r));

  for (let r = 2; r <= rows.length + 1; r++) {
    ws.getCell(`D${r}`).numFmt = '0%';
    ws.getCell(`E${r}`).numFmt = 'yyyy-mm-dd';
    ws.getCell(`F${r}`).numFmt = 'yyyy-mm-dd';
    ws.getCell(`G${r}`).numFmt = '"$"#,##0';
    ws.getCell(`H${r}`).numFmt = '"$"#,##0';
    const status = ws.getCell(`C${r}`).value;
    const fill = {
      'On Track': 'FFD1F2E1',
      'At Risk': 'FFFFE9C7',
      Blocked: 'FFFBE3E3',
      Done: 'FFE5E7EB',
    }[status];
    if (fill) {
      ws.getCell(`C${r}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: fill },
      };
      ws.getCell(`C${r}`).alignment = { horizontal: 'center' };
    }
    ws.getCell(`A${r}`).font = { bold: true };
  }
  zebraDataRows(ws, 2, rows.length + 1, 'FFF5F3FF');
  await saveWorkbook(wb, 'project-tracker.xlsx');
}

async function sprintPlanner() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Sprint Planner';
  const ws = wb.addWorksheet('Sprint', {
    properties: { tabColor: { argb: 'FFDB2777' } },
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'id', width: 8 },
    { key: 'story', width: 38 },
    { key: 'assignee', width: 14 },
    { key: 'points', width: 10 },
    { key: 'status', width: 14 },
    { key: 'epic', width: 18 },
  ]);
  ws.addRow(['ID', 'Story', 'Assignee', 'Points', 'Status', 'Epic']);
  applyHeader(ws.getRow(1), 'FFDB2777');
  const rows = [
    ['CS-101', 'Workbook autosave persists across crashes', 'Aria', 5, 'Done', 'Reliability'],
    ['CS-102', 'Co-edit cursor on frozen panes', 'Liam', 3, 'Done', 'Co-edit'],
    ['CS-103', 'Paste from Numbers preserves currency', 'Maya', 2, 'In Progress', 'xlsx'],
    ['CS-104', 'Home page template gallery', 'Aria', 8, 'In Progress', 'Onboarding'],
    ['CS-105', 'Conditional format icon sets', 'Sam', 5, 'To Do', 'Polish'],
    ['CS-106', 'Pivot drill-down ctrl+shift+d', 'Priya', 3, 'Done', 'Pivots'],
    ['CS-107', 'Chart trendline labels', 'Noah', 2, 'To Do', 'Charts'],
    ['CS-108', 'Goal-seek converges on flat curves', 'Maya', 5, 'To Do', 'Analysis'],
    ['CS-109', 'Sparkline win-loss colors', 'Jordan', 1, 'Done', 'Charts'],
  ];
  rows.forEach((r) => ws.addRow(r));
  for (let r = 2; r <= rows.length + 1; r++) {
    const status = ws.getCell(`E${r}`).value;
    const fill = {
      Done: 'FFD1F2E1',
      'In Progress': 'FFFFE9C7',
      'To Do': 'FFFCE7F3',
    }[status];
    if (fill) {
      ws.getCell(`E${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
      ws.getCell(`E${r}`).alignment = { horizontal: 'center' };
    }
    ws.getCell(`A${r}`).font = { name: 'Calibri', size: 11, color: { argb: 'FF6B7280' } };
  }
  // Totals
  const tr = rows.length + 3;
  ws.getCell(`C${tr}`).value = 'Total points';
  ws.getCell(`D${tr}`).value = { formula: `SUM(D2:D${rows.length + 1})` };
  ws.getCell(`C${tr}`).font = { bold: true };
  ws.getCell(`D${tr}`).font = { bold: true };
  zebraDataRows(ws, 2, rows.length + 1, 'FFFDF2F8');
  await saveWorkbook(wb, 'sprint-planner.xlsx');
}

async function invoice() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Invoice';
  const ws = wb.addWorksheet('Invoice', {
    properties: { tabColor: { argb: 'FFB45309' } },
  });
  setColumns(ws, [
    { key: 'a', width: 30 },
    { key: 'b', width: 14 },
    { key: 'c', width: 14 },
    { key: 'd', width: 16 },
  ]);
  // Header band
  ws.mergeCells('A1:D1');
  ws.getCell('A1').value = 'INVOICE';
  ws.getCell('A1').font = { name: 'Calibri', size: 28, bold: true, color: { argb: 'FFB45309' } };
  ws.getCell('A1').alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(1).height = 44;

  ws.getCell('A3').value = 'Bill to';
  ws.getCell('A3').font = { bold: true, color: { argb: 'FF6B7280' } };
  ws.getCell('A4').value = 'Acme, Inc.';
  ws.getCell('A5').value = '123 Market St';
  ws.getCell('A6').value = 'San Francisco, CA 94103';

  ws.getCell('C3').value = 'Invoice #';
  ws.getCell('D3').value = 'INV-0042';
  ws.getCell('C4').value = 'Date';
  ws.getCell('D4').value = new Date();
  ws.getCell('D4').numFmt = 'yyyy-mm-dd';
  ws.getCell('C5').value = 'Due';
  ws.getCell('D5').value = new Date(Date.now() + 14 * 86400000);
  ws.getCell('D5').numFmt = 'yyyy-mm-dd';
  ws.getCell('C3').font = { bold: true, color: { argb: 'FF6B7280' } };
  ws.getCell('C4').font = { bold: true, color: { argb: 'FF6B7280' } };
  ws.getCell('C5').font = { bold: true, color: { argb: 'FF6B7280' } };

  ws.getRow(8).values = ['Description', 'Qty', 'Rate', 'Amount'];
  applyHeader(ws.getRow(8), 'FFB45309');

  const items = [
    ['Design system audit', 1, 2400],
    ['Brand workshop (2 days)', 2, 1800],
    ['Logo + visual identity', 1, 3200],
    ['Component library setup', 1, 2800],
  ];
  items.forEach((it, i) => {
    const r = 9 + i;
    ws.getCell(`A${r}`).value = it[0];
    ws.getCell(`B${r}`).value = it[1];
    ws.getCell(`B${r}`).alignment = { horizontal: 'center' };
    ws.getCell(`C${r}`).value = it[2];
    ws.getCell(`C${r}`).numFmt = '"$"#,##0.00';
    ws.getCell(`D${r}`).value = { formula: `B${r}*C${r}` };
    ws.getCell(`D${r}`).numFmt = '"$"#,##0.00';
  });
  zebraDataRows(ws, 9, 9 + items.length - 1, 'FFFEF3E7');

  const subRow = 9 + items.length + 1;
  ws.getCell(`C${subRow}`).value = 'Subtotal';
  ws.getCell(`D${subRow}`).value = { formula: `SUM(D9:D${9 + items.length - 1})` };
  ws.getCell(`C${subRow + 1}`).value = 'Tax (8.5%)';
  ws.getCell(`D${subRow + 1}`).value = { formula: `D${subRow}*0.085` };
  ws.getCell(`C${subRow + 2}`).value = 'Total';
  ws.getCell(`D${subRow + 2}`).value = { formula: `D${subRow}+D${subRow + 1}` };
  for (let r = subRow; r <= subRow + 2; r++) {
    ws.getCell(`C${r}`).font = { bold: true };
    ws.getCell(`D${r}`).font = { bold: true };
    ws.getCell(`D${r}`).numFmt = '"$"#,##0.00';
  }
  ws.getCell(`C${subRow + 2}`).font = { bold: true, size: 14, color: { argb: 'FFB45309' } };
  ws.getCell(`D${subRow + 2}`).font = { bold: true, size: 14, color: { argb: 'FFB45309' } };

  await saveWorkbook(wb, 'invoice.xlsx');
}

async function inventory() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Inventory';
  const ws = wb.addWorksheet('Inventory', {
    properties: { tabColor: { argb: 'FF0F766E' } },
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'sku', width: 12 },
    { key: 'name', width: 28 },
    { key: 'category', width: 14 },
    { key: 'qty', width: 10 },
    { key: 'price', width: 12 },
    { key: 'value', width: 14 },
    { key: 'reorder', width: 14 },
  ]);
  ws.addRow(['SKU', 'Name', 'Category', 'Qty', 'Price', 'Value', 'Reorder at']);
  applyHeader(ws.getRow(1), 'FF0F766E');
  const rows = [
    ['SKU-001', 'Espresso beans 1 kg', 'Coffee', 28, 18.5, null, 10],
    ['SKU-002', 'Oat milk 1 L', 'Dairy alt', 64, 3.2, null, 20],
    ['SKU-003', 'Paper cups 12 oz', 'Supplies', 410, 0.08, null, 100],
    ['SKU-004', 'Lids 12 oz', 'Supplies', 380, 0.04, null, 100],
    ['SKU-005', 'Vanilla syrup', 'Syrups', 9, 9.0, null, 6],
    ['SKU-006', 'Chocolate powder', 'Syrups', 5, 12.4, null, 4],
    ['SKU-007', 'Coffee filters', 'Supplies', 80, 0.15, null, 40],
    ['SKU-008', 'Sugar sachets', 'Sweeteners', 1200, 0.01, null, 400],
    ['SKU-009', 'Cleaning spray', 'Cleaning', 6, 4.8, null, 3],
    ['SKU-010', 'Loyalty cards', 'Marketing', 220, 0.05, null, 100],
  ];
  rows.forEach((r) => ws.addRow(r));
  for (let r = 2; r <= rows.length + 1; r++) {
    ws.getCell(`E${r}`).numFmt = '"$"#,##0.00';
    ws.getCell(`F${r}`).value = { formula: `D${r}*E${r}` };
    ws.getCell(`F${r}`).numFmt = '"$"#,##0.00';
    // Reorder-warning style
    const qty = ws.getCell(`D${r}`).value;
    const reorder = ws.getCell(`G${r}`).value;
    if (typeof qty === 'number' && typeof reorder === 'number' && qty <= reorder) {
      ws.getCell(`D${r}`).font = { bold: true, color: { argb: 'FFB91C1C' } };
    }
  }
  zebraDataRows(ws, 2, rows.length + 1, 'FFE6F4F1');
  await saveWorkbook(wb, 'inventory.xlsx');
}

async function expenseReport() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Expense Report';
  const ws = wb.addWorksheet('Expenses', {
    properties: { tabColor: { argb: 'FF0EA5E9' } },
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'date', width: 14 },
    { key: 'merchant', width: 24 },
    { key: 'category', width: 18 },
    { key: 'amount', width: 14 },
    { key: 'reimb', width: 14 },
    { key: 'notes', width: 26 },
  ]);
  ws.addRow(['Date', 'Merchant', 'Category', 'Amount', 'Reimbursable', 'Notes']);
  applyHeader(ws.getRow(1), 'FF0EA5E9');
  const today = new Date();
  const d = (offset) => new Date(today.getTime() + offset * 86400000);
  const rows = [
    [d(-1), 'Blue Bottle Coffee', 'Meals', 8.5, 'Yes', 'Client meeting'],
    [d(-2), 'United Airlines', 'Travel', 412.0, 'Yes', 'SFO → NYC'],
    [d(-3), 'Hilton Midtown', 'Lodging', 248.0, 'Yes', 'NYC offsite'],
    [d(-3), 'Uber', 'Transport', 22.4, 'Yes', 'Airport → hotel'],
    [d(-4), 'Notion subscription', 'Software', 16.0, 'No', 'Personal plan'],
    [d(-5), 'Shake Shack', 'Meals', 18.2, 'Yes', 'Team dinner'],
    [d(-6), 'Office Depot', 'Supplies', 38.95, 'Yes', 'Sticky notes, pens'],
    [d(-10), 'Lyft', 'Transport', 14.3, 'Yes', 'Client onsite'],
  ];
  rows.forEach((r) => ws.addRow(r));
  for (let r = 2; r <= rows.length + 1; r++) {
    ws.getCell(`A${r}`).numFmt = 'yyyy-mm-dd';
    ws.getCell(`D${r}`).numFmt = '"$"#,##0.00';
    ws.getCell(`E${r}`).alignment = { horizontal: 'center' };
    if (ws.getCell(`E${r}`).value === 'Yes')
      ws.getCell(`E${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1F2E1' } };
  }
  zebraDataRows(ws, 2, rows.length + 1, 'FFE0F2FE');
  // Totals
  const tr = rows.length + 3;
  ws.getCell(`C${tr}`).value = 'Total';
  ws.getCell(`D${tr}`).value = { formula: `SUM(D2:D${rows.length + 1})` };
  ws.getCell(`C${tr}`).font = { bold: true };
  ws.getCell(`D${tr}`).font = { bold: true };
  ws.getCell(`D${tr}`).numFmt = '"$"#,##0.00';
  await saveWorkbook(wb, 'expense-report.xlsx');
}

async function classSchedule() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Class Schedule';
  const ws = wb.addWorksheet('Schedule', {
    properties: { tabColor: { argb: 'FFCA8A04' } },
    views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'time', width: 14 },
    { key: 'mon', width: 18 },
    { key: 'tue', width: 18 },
    { key: 'wed', width: 18 },
    { key: 'thu', width: 18 },
    { key: 'fri', width: 18 },
  ]);
  ws.addRow(['Time', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri']);
  applyHeader(ws.getRow(1), 'FFCA8A04');
  const slots = [
    ['08:30 – 09:50', 'Calculus I', '', 'Calculus I', '', 'Calculus I'],
    ['10:00 – 11:20', 'History 101', 'Physics Lab', 'History 101', 'Physics Lab', 'History 101'],
    ['11:30 – 12:50', '', 'Writing Seminar', '', 'Writing Seminar', ''],
    ['13:00 – 14:20', 'CS 250', 'CS 250 — Recitation', 'CS 250', 'CS 250 — Recitation', 'CS 250'],
    ['14:30 – 15:50', 'Office hours', 'Study group', 'Office hours', 'Study group', 'Office hours'],
    ['16:00 – 17:20', '', '', 'Intramural soccer', '', ''],
  ];
  slots.forEach((s) => ws.addRow(s));
  for (let r = 2; r <= slots.length + 1; r++) {
    ws.getCell(`A${r}`).font = { bold: true, color: { argb: 'FF374151' } };
    for (let c = 2; c <= 6; c++) {
      const v = ws.getRow(r).getCell(c).value;
      if (v) {
        ws.getRow(r).getCell(c).alignment = { wrapText: true, vertical: 'middle' };
        ws.getRow(r).getCell(c).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFEF3C7' },
        };
      }
    }
    ws.getRow(r).height = 28;
  }
  await saveWorkbook(wb, 'class-schedule.xlsx');
}

async function gradeTracker() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Grade Tracker';
  const ws = wb.addWorksheet('Grades', {
    properties: { tabColor: { argb: 'FF7C3AED' } },
    views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'student', width: 22 },
    { key: 'hw1', width: 8 },
    { key: 'hw2', width: 8 },
    { key: 'hw3', width: 8 },
    { key: 'mid', width: 10 },
    { key: 'final', width: 10 },
    { key: 'avg', width: 12 },
    { key: 'letter', width: 10 },
  ]);
  ws.addRow(['Student', 'HW1', 'HW2', 'HW3', 'Midterm', 'Final', 'Average', 'Letter']);
  applyHeader(ws.getRow(1), 'FF7C3AED');
  const rows = [
    ['Alice Park', 92, 88, 95, 89, 93],
    ['Ben Singh', 78, 82, 75, 80, 79],
    ['Carla Diaz', 96, 95, 92, 97, 98],
    ['Daniel Wu', 88, 85, 90, 84, 87],
    ['Emily Rao', 70, 72, 78, 74, 71],
    ['Felix Bauer', 90, 92, 88, 91, 95],
    ['Grace Kim', 84, 80, 86, 78, 82],
    ['Henry Cole', 65, 70, 72, 68, 74],
  ];
  rows.forEach((r) => ws.addRow(r));
  for (let r = 2; r <= rows.length + 1; r++) {
    ws.getCell(`G${r}`).value = { formula: `AVERAGE(B${r}:F${r})` };
    ws.getCell(`G${r}`).numFmt = '0.0';
    ws.getCell(`H${r}`).value = {
      formula: `IF(G${r}>=90,"A",IF(G${r}>=80,"B",IF(G${r}>=70,"C",IF(G${r}>=60,"D","F"))))`,
    };
    ws.getCell(`H${r}`).alignment = { horizontal: 'center' };
    ws.getCell(`H${r}`).font = { bold: true };
  }
  zebraDataRows(ws, 2, rows.length + 1, 'FFEDE9FE');
  await saveWorkbook(wb, 'grade-tracker.xlsx');
}

async function travelPlanner() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Travel Planner';
  const ws = wb.addWorksheet('Itinerary', {
    properties: { tabColor: { argb: 'FFE11D48' } },
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'day', width: 12 },
    { key: 'date', width: 14 },
    { key: 'where', width: 22 },
    { key: 'activity', width: 32 },
    { key: 'cost', width: 12 },
    { key: 'notes', width: 30 },
  ]);
  ws.addRow(['Day', 'Date', 'Where', 'Activity', 'Cost', 'Notes']);
  applyHeader(ws.getRow(1), 'FFE11D48');
  const today = new Date();
  const d = (offset) => new Date(today.getTime() + offset * 86400000);
  const rows = [
    [1, d(7), 'Tokyo', 'Arrive at Haneda, settle into hotel', 0, 'JAL flight 745'],
    [2, d(8), 'Tokyo', 'Tsukiji breakfast tour', 65, 'Booked Klook'],
    [2, d(8), 'Tokyo', 'Senso-ji temple + Asakusa walk', 0, ''],
    [3, d(9), 'Tokyo', 'TeamLab Planets', 36, 'Buy ticket online'],
    [4, d(10), 'Hakone', 'Shinkansen → Hakone, onsen', 280, 'Ryokan: Yamayuki'],
    [5, d(11), 'Hakone', 'Lake Ashi pirate ship', 22, ''],
    [6, d(12), 'Kyoto', 'Bullet train → Kyoto', 90, 'Reserve seats'],
    [7, d(13), 'Kyoto', 'Fushimi Inari at sunrise', 0, 'Leave hotel 5am'],
    [8, d(14), 'Osaka', 'Day trip — food crawl', 80, 'Dotonbori'],
  ];
  rows.forEach((r) => ws.addRow(r));
  for (let r = 2; r <= rows.length + 1; r++) {
    ws.getCell(`B${r}`).numFmt = 'yyyy-mm-dd';
    ws.getCell(`E${r}`).numFmt = '"$"#,##0';
  }
  // Total cost
  const tr = rows.length + 3;
  ws.getCell(`D${tr}`).value = 'Total budget';
  ws.getCell(`E${tr}`).value = { formula: `SUM(E2:E${rows.length + 1})` };
  ws.getCell(`D${tr}`).font = { bold: true };
  ws.getCell(`E${tr}`).font = { bold: true };
  ws.getCell(`E${tr}`).numFmt = '"$"#,##0';
  zebraDataRows(ws, 2, rows.length + 1, 'FFFEE2E2');
  await saveWorkbook(wb, 'travel-planner.xlsx');
}

async function meetingNotes() {
  const wb = new ExcelJS.Workbook();
  wb.title = 'Meeting Notes';
  const ws = wb.addWorksheet('Notes', {
    properties: { tabColor: { argb: 'FF334155' } },
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  setColumns(ws, [
    { key: 'date', width: 14 },
    { key: 'mtg', width: 24 },
    { key: 'topic', width: 32 },
    { key: 'decision', width: 32 },
    { key: 'owner', width: 14 },
    { key: 'due', width: 14 },
  ]);
  ws.addRow(['Date', 'Meeting', 'Topic', 'Decision / Action', 'Owner', 'Due']);
  applyHeader(ws.getRow(1), 'FF334155');
  const today = new Date();
  const d = (offset) => new Date(today.getTime() + offset * 86400000);
  const rows = [
    [d(-1), 'Eng weekly', 'Home page', 'Ship behind a feature flag — pick this Friday', 'Aria', d(2)],
    [d(-1), 'Eng weekly', 'xlsx pivots', 'Defer pivot-cache passthrough to P6.1', 'Maya', d(14)],
    [d(-2), 'Design review', 'Template thumbnails', 'Hand-design > auto-render — better artistic control', 'Liam', d(3)],
    [d(-3), '1:1 — Sam', 'Search v2 launch', 'Soft launch next Wed; full launch following Mon', 'Sam', d(10)],
    [d(-5), 'All hands', 'Mobile redesign', 'Will defer until Q3 if pricing PR slips', 'Liam', d(14)],
    [d(-7), 'Eng weekly', 'Yjs upgrade', '0.18.x → 0.19.x — schedule for slow week', 'Priya', d(30)],
  ];
  rows.forEach((r) => ws.addRow(r));
  for (let r = 2; r <= rows.length + 1; r++) {
    ws.getCell(`A${r}`).numFmt = 'yyyy-mm-dd';
    ws.getCell(`F${r}`).numFmt = 'yyyy-mm-dd';
    ws.getCell(`E${r}`).alignment = { horizontal: 'center' };
    ws.getRow(r).height = 32;
    ws.getRow(r).alignment = { wrapText: true, vertical: 'top' };
  }
  zebraDataRows(ws, 2, rows.length + 1, 'FFF1F5F9');
  await saveWorkbook(wb, 'meeting-notes.xlsx');
}

// ── Run ─────────────────────────────────────────────────────────────────────

console.info(`Building templates into ${OUT_DIR}`);
await personalBudget();
await todoList();
await projectTracker();
await sprintPlanner();
await invoice();
await inventory();
await expenseReport();
await classSchedule();
await gradeTracker();
await travelPlanner();
await meetingNotes();
console.info('Done. 11 templates generated.');
