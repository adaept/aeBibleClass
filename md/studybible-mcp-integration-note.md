# Pointer: studybible-mcp integration plan

**This is a pointer, not the plan.** The full analysis lives in the studio hub repo,
`adaept5tudio`, per that repo's convention of keeping cross-project/master plans there:

`C:\adaept\adaept5tudio\docs\studybible-mcp-integration-plan.md`

## What it covers

An evaluation of [github.com/peterennis/studybible-mcp](https://github.com/peterennis/studybible-mcp)
(the user's own fork of `djayatillake/studybible-mcp`, an 18-tool MCP server over a ~600MB
scholarly Bible dataset — Greek/Hebrew lexicons, morphology, study notes, genealogy graphs, Ancient
Near East cultural context) as a data source for this project's Word 365 JS add-in conversion
(see `adaept5tudio/docs/aeBibleClass-word-addin-conversion-plan.md`).

Covers: two integration architectures (direct tool invocation vs. a full agentic chat panel — the
former recommended first), licensing/attribution obligations (STEPBible CC BY 4.0, Aquifer content
CC BY-SA 4.0 — note the share-alike term, stricter than this studio's usual permissive-tier asset
policy), an explicit flag that this project's own **Bias Guard** charter (see this file's parent,
`md/README.md` §"Bias Guard Integration") applies to any content sourced from that server before it
ships in a UI, an i18n gap analysis (the dataset's Strong's numbers/morphology/verse references are
language-agnostic; its scholarly prose — lexicon definitions, study notes, hermeneutical guidance —
is English-only today), and a now-vs-later recommendation.

Read the full plan in `adaept5tudio` before acting on any of this — this note only exists so a
future reader of this repo's own docs knows the plan exists and where to find it.
