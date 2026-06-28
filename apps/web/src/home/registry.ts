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
 * Template library for the home page.
 *
 * Each entry maps 1:1 to a real .xlsx file in `apps/web/public/templates/`
 * (built by `scripts/build-templates.mjs`). The `preview` field is a
 * hand-tuned tiny grid that drives the thumbnail render — fast to draw,
 * no xlsx parse needed before the user actually picks.
 *
 * Clicking a card → fetch /templates/{id}.xlsx → parse through the same
 * worker pipeline File→Open uses → replaceWorkbook. The home overlay
 * dismisses automatically because the workbook is no longer Untitled.
 */

export type CellFormat = 'currency' | 'percent' | 'date' | 'badge' | 'muted' | 'bold';
export type PreviewCell =
  | string
  | { v: string; fmt?: CellFormat; color?: string };
export type TemplatePreview = {
  /** Header labels (always rendered with the accent strip). */
  header: string[];
  /** Body rows. Variable lengths OK — empty cells render as blanks. */
  rows: PreviewCell[][];
};

export type TemplateCategory = 'Personal' | 'Work' | 'Finance' | 'Education';

export type Template = {
  /** Slug + filename stem (`{id}.xlsx`). */
  id: string;
  name: string;
  category: TemplateCategory;
  description: string;
  /** Hex accent — drives the card header strip and thumbnail accent. */
  accent: string;
  /** Material Symbols icon name shown on the card chrome. */
  icon: string;
  /** Featured templates appear in the hero strip. */
  featured?: boolean;
  preview: TemplatePreview;
};

// Each preview is intentionally short — first 4-5 visible rows — so the
// thumbnail reads at-a-glance. Don't bloat with detail the user can't see.

export const TEMPLATES: Template[] = [
  {
    id: 'blank',
    name: 'Blank spreadsheet',
    category: 'Personal',
    description: 'Start fresh with an empty workbook.',
    accent: '#6B7280',
    icon: 'add',
    featured: true,
    preview: {
      header: ['A', 'B', 'C', 'D'],
      rows: [
        ['', '', '', ''],
        ['', '', '', ''],
        ['', '', '', ''],
        ['', '', '', ''],
      ],
    },
  },
  {
    id: 'personal-budget',
    name: 'Personal budget',
    category: 'Finance',
    description: 'Monthly income, expenses, and savings rate with totals.',
    accent: '#2E7D32',
    icon: 'savings',
    featured: true,
    preview: {
      header: ['Category', 'Budget', 'Jan', 'Feb', 'Avg'],
      rows: [
        ['Salary', { v: '$5,400', fmt: 'currency' }, { v: '$5,400', fmt: 'currency' }, { v: '$5,400', fmt: 'currency' }, { v: '$5,400', fmt: 'currency' }],
        ['Rent', { v: '$1,800', fmt: 'currency' }, { v: '$1,800', fmt: 'currency' }, { v: '$1,800', fmt: 'currency' }, { v: '$1,800', fmt: 'currency' }],
        ['Groceries', { v: '$520', fmt: 'currency' }, { v: '$510', fmt: 'currency' }, { v: '$530', fmt: 'currency' }, { v: '$520', fmt: 'currency' }],
        ['Savings', { v: '$800', fmt: 'currency' }, { v: '$800', fmt: 'currency' }, { v: '$800', fmt: 'currency' }, { v: '$800', fmt: 'currency' }],
      ],
    },
  },
  {
    id: 'todo-list',
    name: 'To-do list',
    category: 'Personal',
    description: 'Task list with status, priority, and due dates.',
    accent: '#1D4ED8',
    icon: 'check_circle',
    featured: true,
    preview: {
      header: ['Task', 'Status', 'Priority', 'Due'],
      rows: [
        ['Draft Q3 proposal', { v: 'In Progress', fmt: 'badge', color: '#B45309' }, { v: 'High', fmt: 'bold', color: '#B91C1C' }, { v: 'May 26', fmt: 'date' }],
        ['Reply to Sara', { v: 'To Do', fmt: 'badge', color: '#4F46E5' }, { v: 'High', fmt: 'bold', color: '#B91C1C' }, { v: 'May 25', fmt: 'date' }],
        ['Onboarding deck', { v: 'Done', fmt: 'badge', color: '#15803D' }, 'Medium', { v: 'May 21', fmt: 'date' }],
        ['Book dentist', { v: 'To Do', fmt: 'badge', color: '#4F46E5' }, 'Low', { v: 'Jun 7', fmt: 'date' }],
      ],
    },
  },
  {
    id: 'project-tracker',
    name: 'Project tracker',
    category: 'Work',
    description: 'Project status, progress, owner, budget vs. spent.',
    accent: '#5B21B6',
    icon: 'view_kanban',
    featured: true,
    preview: {
      header: ['Project', 'Owner', 'Status', '%', 'Budget'],
      rows: [
        ['Onboarding revamp', 'Aria', { v: 'On Track', fmt: 'badge', color: '#15803D' }, { v: '60%', fmt: 'percent' }, { v: '$18,000', fmt: 'currency' }],
        ['Mobile redesign', 'Liam', { v: 'At Risk', fmt: 'badge', color: '#B45309' }, { v: '35%', fmt: 'percent' }, { v: '$42,000', fmt: 'currency' }],
        ['Search v2', 'Maya', { v: 'On Track', fmt: 'badge', color: '#15803D' }, { v: '78%', fmt: 'percent' }, { v: '$22,000', fmt: 'currency' }],
        ['Pricing refresh', 'Aria', { v: 'Done', fmt: 'badge', color: '#374151' }, { v: '100%', fmt: 'percent' }, { v: '$8,000', fmt: 'currency' }],
      ],
    },
  },
  {
    id: 'sprint-planner',
    name: 'Sprint planner',
    category: 'Work',
    description: 'Backlog stories with assignee, points, and status.',
    accent: '#DB2777',
    icon: 'rocket_launch',
    preview: {
      header: ['ID', 'Story', 'Assignee', 'Pts', 'Status'],
      rows: [
        [{ v: 'CS-101', fmt: 'muted' }, 'Autosave reliability', 'Aria', '5', { v: 'Done', fmt: 'badge', color: '#15803D' }],
        [{ v: 'CS-102', fmt: 'muted' }, 'Frozen-pane cursors', 'Liam', '3', { v: 'Done', fmt: 'badge', color: '#15803D' }],
        [{ v: 'CS-104', fmt: 'muted' }, 'Home template gallery', 'Aria', '8', { v: 'In Prog', fmt: 'badge', color: '#B45309' }],
        [{ v: 'CS-105', fmt: 'muted' }, 'Conditional formats', 'Sam', '5', { v: 'To Do', fmt: 'badge', color: '#DB2777' }],
      ],
    },
  },
  {
    id: 'invoice',
    name: 'Invoice',
    category: 'Finance',
    description: 'Itemised invoice with subtotal, tax, and total.',
    accent: '#B45309',
    icon: 'receipt_long',
    preview: {
      header: ['Description', 'Qty', 'Rate', 'Amount'],
      rows: [
        ['Design audit', '1', { v: '$2,400', fmt: 'currency' }, { v: '$2,400', fmt: 'currency' }],
        ['Brand workshop', '2', { v: '$1,800', fmt: 'currency' }, { v: '$3,600', fmt: 'currency' }],
        ['Visual identity', '1', { v: '$3,200', fmt: 'currency' }, { v: '$3,200', fmt: 'currency' }],
        [{ v: 'Total', fmt: 'bold' }, '', '', { v: '$10,408', fmt: 'currency' }],
      ],
    },
  },
  {
    id: 'inventory',
    name: 'Inventory',
    category: 'Work',
    description: 'SKU, qty on hand, unit price, and reorder threshold.',
    accent: '#0F766E',
    icon: 'inventory_2',
    preview: {
      header: ['SKU', 'Name', 'Qty', 'Price', 'Value'],
      rows: [
        [{ v: 'SKU-001', fmt: 'muted' }, 'Espresso beans', '28', { v: '$18.50', fmt: 'currency' }, { v: '$518', fmt: 'currency' }],
        [{ v: 'SKU-002', fmt: 'muted' }, 'Oat milk', '64', { v: '$3.20', fmt: 'currency' }, { v: '$205', fmt: 'currency' }],
        [{ v: 'SKU-005', fmt: 'muted' }, 'Vanilla syrup', { v: '9', fmt: 'bold', color: '#B91C1C' }, { v: '$9.00', fmt: 'currency' }, { v: '$81', fmt: 'currency' }],
        [{ v: 'SKU-008', fmt: 'muted' }, 'Sugar sachets', '1,200', { v: '$0.01', fmt: 'currency' }, { v: '$12', fmt: 'currency' }],
      ],
    },
  },
  {
    id: 'expense-report',
    name: 'Expense report',
    category: 'Finance',
    description: 'Itemised expenses for reimbursement.',
    accent: '#0EA5E9',
    icon: 'request_quote',
    preview: {
      header: ['Date', 'Merchant', 'Category', 'Amount'],
      rows: [
        [{ v: 'May 23', fmt: 'date' }, 'Blue Bottle', 'Meals', { v: '$8.50', fmt: 'currency' }],
        [{ v: 'May 22', fmt: 'date' }, 'United', 'Travel', { v: '$412.00', fmt: 'currency' }],
        [{ v: 'May 21', fmt: 'date' }, 'Hilton', 'Lodging', { v: '$248.00', fmt: 'currency' }],
        [{ v: 'May 19', fmt: 'date' }, 'Uber', 'Transport', { v: '$22.40', fmt: 'currency' }],
      ],
    },
  },
  {
    id: 'class-schedule',
    name: 'Class schedule',
    category: 'Education',
    description: 'Weekday columns with time slots and class blocks.',
    accent: '#CA8A04',
    icon: 'school',
    preview: {
      header: ['Time', 'Mon', 'Tue', 'Wed', 'Thu'],
      rows: [
        ['08:30', { v: 'Calculus I', fmt: 'badge', color: '#CA8A04' }, '', { v: 'Calculus I', fmt: 'badge', color: '#CA8A04' }, ''],
        ['10:00', { v: 'History 101', fmt: 'badge', color: '#CA8A04' }, { v: 'Physics Lab', fmt: 'badge', color: '#CA8A04' }, { v: 'History 101', fmt: 'badge', color: '#CA8A04' }, { v: 'Physics Lab', fmt: 'badge', color: '#CA8A04' }],
        ['13:00', { v: 'CS 250', fmt: 'badge', color: '#CA8A04' }, { v: 'CS 250 Rec', fmt: 'badge', color: '#CA8A04' }, { v: 'CS 250', fmt: 'badge', color: '#CA8A04' }, { v: 'CS 250 Rec', fmt: 'badge', color: '#CA8A04' }],
        ['14:30', 'Office hrs', 'Study grp', 'Office hrs', 'Study grp'],
      ],
    },
  },
  {
    id: 'grade-tracker',
    name: 'Grade tracker',
    category: 'Education',
    description: 'Students, assignments, weighted average, letter grade.',
    accent: '#7C3AED',
    icon: 'grade',
    preview: {
      header: ['Student', 'HW1', 'Mid', 'Final', 'Avg', 'Letter'],
      rows: [
        ['Alice Park', '92', '89', '93', '91.4', { v: 'A', fmt: 'bold', color: '#15803D' }],
        ['Ben Singh', '78', '80', '79', '78.8', { v: 'C', fmt: 'bold', color: '#B45309' }],
        ['Carla Diaz', '96', '97', '98', '95.6', { v: 'A', fmt: 'bold', color: '#15803D' }],
        ['Daniel Wu', '88', '84', '87', '86.8', { v: 'B', fmt: 'bold', color: '#1D4ED8' }],
      ],
    },
  },
  {
    id: 'travel-planner',
    name: 'Travel planner',
    category: 'Personal',
    description: 'Day-by-day itinerary with locations and costs.',
    accent: '#E11D48',
    icon: 'flight_takeoff',
    preview: {
      header: ['Day', 'Date', 'Where', 'Activity', 'Cost'],
      rows: [
        ['1', { v: 'May 30', fmt: 'date' }, 'Tokyo', 'Arrive Haneda', { v: '$0', fmt: 'currency' }],
        ['2', { v: 'May 31', fmt: 'date' }, 'Tokyo', 'Tsukiji breakfast', { v: '$65', fmt: 'currency' }],
        ['4', { v: 'Jun 2', fmt: 'date' }, 'Hakone', 'Onsen + ryokan', { v: '$280', fmt: 'currency' }],
        ['6', { v: 'Jun 4', fmt: 'date' }, 'Kyoto', 'Bullet train', { v: '$90', fmt: 'currency' }],
      ],
    },
  },
  {
    id: 'meeting-notes',
    name: 'Meeting notes',
    category: 'Work',
    description: 'Running decisions and action items from meetings.',
    accent: '#334155',
    icon: 'event_note',
    preview: {
      header: ['Date', 'Meeting', 'Decision / Action', 'Owner'],
      rows: [
        [{ v: 'May 23', fmt: 'date' }, 'Eng weekly', 'Ship home page Friday', 'Aria'],
        [{ v: 'May 23', fmt: 'date' }, 'Eng weekly', 'Defer pivot cache', 'Maya'],
        [{ v: 'May 22', fmt: 'date' }, 'Design review', 'Hand-design thumbnails', 'Liam'],
        [{ v: 'May 21', fmt: 'date' }, '1:1 — Sam', 'Search v2 soft launch', 'Sam'],
      ],
    },
  },
];

export const CATEGORIES: TemplateCategory[] = [
  'Personal',
  'Work',
  'Finance',
  'Education',
];
