# Route A notes (slide-beautify, maintainers only; not shipped)

Things to know before changing the skill's writer or catalogue. Kept short; each line says why.

- **Re-run the XML comparison when the catalogue grows.** The Reviewer compared every shape the skill writes (547 on the 26 sample pages) against the library's own writer (ooxml.js). The sample has only simple shapes. When the catalogue gains tables, groups, pictures or quarter-turned shapes, run that comparison again before shipping (Reviewer, 14 Nov 2026).
- **Spacing a hair off the library size.** A deck a hair off 13.333 x 7.5 in (13.33 in wide, or a title ending 0.001 to 0.05 in lower than the library's) gets paragraph spacing half a point under the library's (11.5 for 12): Place.spacing rounds down to half points. Safe side and not visible; the CEO decided to leave it (14 Nov 2026). Do not add a "scale of 0.995 counts as 1" rule.
- **Adjust values come from a --keep-adj export.** pages.json is built from `node tools/export-gslides-library.mjs --out <dir> --keep-adj` (lavender). The plain export (what the Google Slides add-on serves) leaves them out, and pills and chevrons then lose their corners.
- **Line figures are provisional** for Aptos and Aptos Display (1.285, the Windows figure) until the Windows render reads them (WC render decks, 14 Nov 2026).
- **Options off by default:** `--fit-first` and `--label-half-step` wait on the CEO's word (and Davon's pick for the label step). Their sheets are in the CEO briefs.
