#!/usr/bin/env python3
# Copyright 2026 Casual Office
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""
Subset the Material Symbols Outlined variable font down to ONLY the icon
ligatures the sheets app actually renders.

Why: icons in this app are font ligatures (apps/web/src/shell/Icon.tsx renders
the icon name as ligature text). The full variable font is ~3.8 MB woff2; the
app uses ~150 distinct glyphs. Subsetting drops it to <100 KB while keeping all
four variable axes (FILL / GRAD / opsz / wght) so filled/weighted icons still
shape correctly.

Pipeline:
  1. Collect every icon name used across apps/web/src — static `<Icon name="x">`
     literals, `icon: 'x'` config keys, ternary strings inside `<Icon name={...}>`,
     plus the NAME_REMAP targets in Icon.tsx (e.g. `ink_highlighter`).
  2. Shape each name through the MASTER font (HarfBuzz). A real Material Symbols
     icon collapses its whole name to exactly ONE glyph; anything that doesn't is
     a false positive (e.g. a radio `name="orientation"`) and is dropped.
  3. Subset by the resolved glyph ids + the 28 component letter/digit/underscore
     glyphs, with --no-layout-closure so the ligature table is pruned to ONLY the
     wanted ligatures (a plain --text subset keeps the whole 6000-icon ligature
     set because every icon's input chars are present).
  4. Re-verify: every collected name must still shape to exactly one glyph in the
     subset. If not, the script exits non-zero and writes nothing.

Requires: fontTools, uharfbuzz, and a HarfBuzz with the woff2 master decoded to
TTF first (HarfBuzz here lacks brotli, so we decompress in-process).

Usage:
  python3 scripts/fonts/subset-material-symbols.py \
      --master /path/to/material-symbols-outlined.woff2 \
      --out apps/web/public/fonts/material-symbols-outlined.subset.woff2
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
SRC = os.path.join(REPO, "apps", "web", "src")

# NAME_REMAP targets in apps/web/src/shell/Icon.tsx — the rendered ligature when
# the named input differs from what is drawn. Keep both input + output names.
EXTRA_NAMES = {"ink_highlighter", "format_color_fill"}

PAT_QUOTED = re.compile(r"""['"]([a-z0-9_]+)['"]""")
# `icon: 'x'` object key (ribbon/menu/panel config data) and `icon="x"` JSX prop
PAT_ICON_KEY = re.compile(r"""\bicon:\s*["']([a-z0-9_]+)["']""")
PAT_ICON_PROP = re.compile(r"""\bicon=["']([a-z0-9_]+)["']""")
# `icon={ ... 'x' : 'y' }` JSX prop with a ternary (e.g. merge/unmerge swap)
PAT_ICON_PROP_EXPR = re.compile(r"""\bicon=\{([^}]*)\}""")
# `<Icon ... name="x">` (multiline-safe) and `<Icon name={ ... 'x' : 'y' }>`
PAT_ICON_TAG_NAME = re.compile(r"""<Icon\b[^>]*?\bname=["']([a-z0-9_]+)["']""", re.S)
PAT_ICON_TAG_EXPR = re.compile(r"""<Icon\b[^>]*?\bname=\{([^}]*)\}""", re.S)
# bare `name="x"` — broad; the master-font shaping pass below filters out
# non-icons (radio input names, enum values, …) by requiring a 1-glyph ligature.
PAT_NAME_ATTR = re.compile(r"""name=["']([a-z0-9_]+)["']""")


def collect_candidate_names():
    names = set(EXTRA_NAMES)
    for root, _dirs, files in os.walk(SRC):
        for f in files:
            if not f.endswith((".ts", ".tsx")):
                continue
            txt = open(os.path.join(root, f), encoding="utf-8").read()
            for m in PAT_ICON_KEY.finditer(txt):
                names.add(m.group(1))
            for m in PAT_ICON_PROP.finditer(txt):
                names.add(m.group(1))
            for m in PAT_ICON_PROP_EXPR.finditer(txt):
                for q in PAT_QUOTED.finditer(m.group(1)):
                    names.add(q.group(1))
            for m in PAT_ICON_TAG_NAME.finditer(txt):
                names.add(m.group(1))
            for m in PAT_ICON_TAG_EXPR.finditer(txt):
                for q in PAT_QUOTED.finditer(m.group(1)):
                    names.add(q.group(1))
            for m in PAT_NAME_ATTR.finditer(txt):
                names.add(m.group(1))
    return names


def to_ttf(woff2_path, out_ttf):
    from fontTools.ttLib import TTFont

    f = TTFont(woff2_path)
    f.flavor = None
    f.save(out_ttf)


def shape_map(ttf_path, names):
    """Return {name: [glyph ids]} for each name shaped through the font."""
    import uharfbuzz as hb

    blob = hb.Blob.from_file_path(ttf_path)
    face = hb.Face(blob)
    font = hb.Font(face)
    hb.ot_font_set_funcs(font)
    out = {}
    for n in sorted(names):
        buf = hb.Buffer()
        buf.add_str(n)
        buf.guess_segment_properties()
        hb.shape(font, buf, {})  # rlig/rclt are default-on
        out[n] = [g.codepoint for g in buf.glyph_infos]
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--master", required=True, help="path to full material-symbols woff2")
    ap.add_argument("--out", required=True, help="output subset woff2 path")
    ap.add_argument("--names-out", default=os.path.join(os.path.dirname(__file__), "icon-names.json"))
    args = ap.parse_args()

    with tempfile.TemporaryDirectory() as td:
        master_ttf = os.path.join(td, "master.ttf")
        to_ttf(args.master, master_ttf)

        from fontTools.ttLib import TTFont

        candidates = collect_candidate_names()
        shaped = shape_map(master_ttf, candidates)

        icons = {}
        rejected = {}
        for n, gids in shaped.items():
            if len(gids) == 1 and gids[0] != 0:
                icons[n] = gids[0]
            else:
                rejected[n] = gids
        if rejected:
            print("Dropped non-icon candidates (do not ligate to one glyph):")
            for n, g in sorted(rejected.items()):
                print(f"  - {n}: {g}")

        names = sorted(icons)
        json.dump(names, open(args.names_out, "w"), indent=1)
        print(f"Collected {len(names)} real icon names -> {args.names_out}")

        ttf = TTFont(master_ttf)
        go = ttf.getGlyphOrder()
        cmap = ttf.getBestCmap()
        chars = sorted(set("".join(names)))
        comp_gids = {go.index(cmap[ord(c)]) for c in chars}
        all_gids = sorted(set(icons.values()) | comp_gids)
        gids_file = os.path.join(td, "gids.txt")
        open(gids_file, "w").write(",".join(str(g) for g in all_gids))

        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        cmd = [
            "pyftsubset", args.master,
            f"--output-file={args.out}",
            "--flavor=woff2",
            f"--gids-file={gids_file}",
            "--layout-features=rlig,rclt,liga,calt,dlig,ccmp,locl,mark,markmk",
            "--no-layout-closure",
            "--no-hinting",
            "--desubroutinize",
        ]
        print("Running:", " ".join(cmd))
        subprocess.run(cmd, check=True)

        # Re-verify coverage on the produced subset.
        subset_ttf = os.path.join(td, "subset.ttf")
        to_ttf(args.out, subset_ttf)
        recheck = shape_map(subset_ttf, set(names))
        bad = {n: g for n, g in recheck.items() if len(g) != 1 or g[0] == 0}
        if bad:
            print("COVERAGE FAILED on subset:")
            for n, g in sorted(bad.items()):
                print(f"  - {n}: {g}")
            os.unlink(args.out)
            sys.exit(1)

        master_sz = os.path.getsize(args.master)
        out_sz = os.path.getsize(args.out)
        print(f"COVERAGE PASS: {len(names)}/{len(names)} icons shape to a single glyph.")
        print(f"Size: {master_sz/1e6:.2f} MB -> {out_sz/1024:.1f} KB "
              f"({100*(1-out_sz/master_sz):.1f}% smaller)")


if __name__ == "__main__":
    main()
