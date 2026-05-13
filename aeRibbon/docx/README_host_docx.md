# aeRibbon-host.docx — manual creation

`aeRibbon-host.docx` is a **manually-authored** empty Word document used
only for Gate G7 (template install/load smoke). The Write tool cannot
produce binary `.docx` files, so the Editor/Developer creates it once:

1. Open Word 365 → File → New → Blank document.
2. Paste this single paragraph as the only content:

   > Attach `aeRibbon.dotm` via File > Options > Add-ins (Manage:
   > Templates → Go → Add), then open a Radiant Word Bible `.docx`
   > to see the **Radiant Word Bible** ribbon tab.

3. File → Save As → **Word Document (`*.docx`)** → save here as
   `aeRibbon-host.docx`.
4. Close Word.

The file must contain **no Bible text** and **no macros**. Its only job
is to prove the template loads and the ribbon tab renders without errors
(Gate G7). Real navigation testing happens against the production Bible
`.docx` (Gate G8).

The production Bible `.docx` itself is also produced manually — see the
"Producing the production Bible `.docx`" section in `aeRibbon/BUILD.md`.
