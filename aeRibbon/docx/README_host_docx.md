# aeRibbon-host.docx — manual creation

`aeRibbon-host.docx` is a **manually-authored** empty Word document used
only for Gate G7 (template install/load smoke). The Write tool cannot
produce binary `.docx` files, so the Editor/Developer creates it once:

1. Open Word 365 → File → New → Blank document.
2. Paste this single paragraph as the only content — deliberately
   generic, with no specific filename baked in (the production docx's
   name has already changed once; this fixture shouldn't need
   re-authoring every time it does again):

   > Attach `aeRibbon.dotm` (see `aeRibbon/BUILD.md`'s "Attaching the
   > template"), then open the current production Bible `.docx` to see
   > the **Radiant Word Bible** ribbon tab in context.

3. File → Save As → **Word Document (`*.docx`)** → save here as
   `aeRibbon-host.docx`.
4. Close Word.

The file must contain **no Bible text** and **no macros**. Its only job
is to prove the template loads and the ribbon tab renders without errors
(Gate G7). Real navigation testing happens against the production Bible
`.docx` (Gate G8).

**Use the Startup-folder attach method for this file**, not the
per-document Templates-dialog method the paragraph might otherwise
suggest — see `aeRibbon/BUILD.md`'s "Attaching the template" for why
(global vs. per-document scope). After confirming the ribbon tab
renders, close this file **without saving** — nothing document-specific
should have changed.

**Before opening any production Bible `.docx` (Gate G8), it must already
have been produced via the full procedure** — see the "Producing the
production Bible `.docx`" section in `aeRibbon/BUILD.md`: Save-As from
the dev `.docm`, then `py/strip_ribbon.py`, then the `customUI` guard
must return nothing. Skipping the strip step and opening a fresh Save-As
directly is a known, reliable trigger for "the macro can't be found" (6x)
and the ribbon failing to load — not a new bug, just a step done out of
order.
