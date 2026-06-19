/**
 * Univer plugin CSS — side-effect-only imports. Each plugin ships its own
 * design tokens / layout primitives that must load once per app.
 *
 * Keep this file as a flat list so the bundler can tree-shake unused tokens
 * later if we drop a plugin. Order doesn't matter; cascade is by ID/specificity.
 */
import '@univerjs/design/lib/index.css';
import '@univerjs/ui/lib/index.css';
import '@univerjs/docs-ui/lib/index.css';
import '@univerjs/sheets-ui/lib/index.css';
import '@univerjs/sheets-formula-ui/lib/index.css';
import '@univerjs/sheets-sort-ui/lib/index.css';
import '@univerjs/sheets-filter-ui/lib/index.css';
import '@univerjs/sheets-numfmt-ui/lib/index.css';
import '@univerjs/find-replace/lib/index.css';
import '@univerjs/sheets-conditional-formatting-ui/lib/index.css';
import '@univerjs/sheets-data-validation-ui/lib/index.css';
import '@univerjs/sheets-hyper-link-ui/lib/index.css';
import '@univerjs/sheets-note-ui/lib/index.css';
import '@univerjs/sheets-table-ui/lib/index.css';
import '@univerjs/sheets-thread-comment-ui/lib/index.css';
import '@univerjs/thread-comment-ui/lib/index.css';
import '@univerjs/drawing-ui/lib/index.css';
import '@univerjs/sheets-drawing-ui/lib/index.css';
import '@univerjs/sheets-crosshair-highlight/lib/index.css';
import '@univerjs/sheets-zen-editor/lib/index.css';
