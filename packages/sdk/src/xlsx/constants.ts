/**
 * Name of a hidden worksheet we use to stash JSON we can't represent
 * natively in xlsx (Univer plugin state — e.g. table definitions, outline
 * groups). On open we recognize and consume it, never showing it to the
 * user. Defined in its own module so both the parser worker and the
 * exporter worker can import it without dragging the other one's
 * ExcelJS code into their bundle.
 */
export const RESOURCES_SHEET = '__casual_sheets_resources__';
