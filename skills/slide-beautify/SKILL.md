---
name: slide-beautify
description: Rebuilds rough or plain slides as clean consulting pages. Reads a .pptx, works out what each slide is (an executive summary, a 2x2, a process, columns of ideas, a list of decisions, a timeline), and lays its own words out again on the best-matching page from a library of designed consulting layouts, in the deck's own fonts and colours. Returns a new file and a list of every change, page by page; the original is never changed. Use when the user asks to beautify, clean up, redesign, re-lay out, make professional, or make consulting-style the slides in a deck, or to turn a wall of bullets into a proper page. It changes layout, not content: to check numbers and sources use slide review; to write the words use the slide writer.
---

# Slide beautify

Take a deck whose content is right but whose pages look rough, and give each page a proper consulting layout without changing what it says. Every rebuilt page comes from a library page drawn in the deck's own theme, so it matches the rest of the deck and stays fully editable in PowerPoint or Google Slides.

## 1. Read the deck

You need the .pptx file and a place to run Python with python-pptx (`pip install python-pptx` if it is missing). Run, from this skill's folder:

```
python3 scripts/read_deck.py "<deck.pptx>" --pages
```

In Claude Code the full path is `${CLAUDE_SKILL_DIR}/scripts/read_deck.py`. The output is JSON:
- `slides`: for each slide its title, its text boxes (paragraphs, levels, bold lead-ins), tables, other objects, source line and notes, and `rebuild` (true, or false with the reason in `why_not`);
- `pages`: the library pages you can use, each with `fits` (what it is for) and its text `slots` (`id`, `role`, the `sample` text it shows, and `max_chars`, roughly how much fits);
- `theme`: the deck's fonts and colours, for your report.

Only text pages can be rebuilt. Slides with charts, pictures, tables, SmartArt, groups or media are left as they are, and the report says so.

No file, only screenshots or pasted text? Say that this skill needs the .pptx, and offer the slide writer or slide style instead.

## 2. Choose a page for each slide

Go slide by slide. Skip the cover, section dividers and any slide marked `rebuild: false`. For each of the rest:

1. **Say what the slide is in one phrase**: "five topics, each a label and a sentence", "four options sorted by impact and ease", "three steps in order".
2. **Pick the page whose `fits` matches it**, and whose number of heading and body slots matches the number of items on the slide. Five items go on a five-row page, not a four-column one. If nothing fits well, leave the slide alone and say why: a forced layout is worse than a plain one.
3. **Never pick a page because it looks impressive.** A list stays a list unless the content really is a 2x2, a process or a comparison.

## 3. Write the plan

Write a JSON file, `plan.json`:

```json
{"pages": [
  {"slide": 3, "page": "executive-summary",
   "title": "The slide's own title",
   "slots": {"s1": "Context", "s2": "The market grew 6% a year since 2021, but our share fell two points"},
   "keep": ["s11"],
   "clear": ["s12"],
   "source": "Source: the slide's own source line"}
]}
```

Rules for the content:
- **Use the slide's own words.** Move them into the slots; shorten only to fit `max_chars`, and never change a number, a name or a claim. Do not add content the slide did not have.
- **Keep the title** as it is. If it is a topic ("Summary") rather than a message, keep it and suggest a message title in your report; change it only if the user asked for that.
- **Headings** (`role: heading`) take short labels: the lead-in of a bullet ("Context:") or a few words. **Body** slots take the rest of that item. Split "Label: text" bullets this way.
- A slot value is a string or a list of paragraphs; `"**Lead-in:** rest"` makes the lead-in bold.
- **Every slot must be filled, kept or cleared.** `keep` leaves the page's own sample text, for generic labels that still fit ("Pros", "Strengths", an axis name like "Impact"). `clear` empties a slot. Slots with `role: fixed` (numbers like "1", "2") look after themselves.
- Never leave sample text that does not belong to this deck. If a page has more slots than the slide has items, choose a smaller page instead of clearing a lot.
- Carry the slide's source line into `source`. Speaker notes are kept automatically.

## 4. Build the new deck

```
python3 scripts/beautify_deck.py "<deck.pptx>" plan.json "<deck> (beautified).pptx" --report report.json
```

It writes a new file. The original is never changed. Each planned slide is replaced in place by its rebuilt page; add `--keep-originals` to keep each original right after it, hidden, for comparison. If the plan has a mistake (an unknown page or slot, a slot left open), the script stops and says what to fix: fix the plan and run it again.

The report lists, for each rebuilt slide, the slots filled, kept and cleared, any text that ended smaller than the page drew it (as the slide carries it, after the layout pass below), and warnings for text that may still overflow. Shorten that text in the plan and run again until there are no overflow warnings.

Then each page gets the library's own layout pass, so it reads full and even with the user's words: every plain text box is fitted to its text (Resize Shape to Fit Text stays on), a row of cards or a list of rows that ends the page is spread to the bottom of the body (never closer than drawn), and body text is raised half a point at a time while everything still fits (9 to 13.5 pt text only, never past 14 pt). The report says what it did under `layout` and how many points the body text went up under `text_up`. Add `--no-fit` to draw the pages exactly as the library drew them, for comparison.

Text is measured with width tables for the deck's theme faces, at the size each text will have on the slide: Arial, Georgia, Verdana, Tahoma, Calibri, Calibri Light, Aptos, Cambria, Century Gothic, Franklin Gothic Book, Tw Cen MT and Times New Roman. A deck whose shape differs from the library's 16:9 (4:3, or a title that ends low) keeps 3% more room. When the deck uses a face without a table (Segoe UI, Montserrat, Poppins, Aptos Display and others), the page is fitted as if the face were 15% wider with taller lines, its text is never enlarged, and the report warns: tell the user to check those pages in PowerPoint. Kinds of text keep one size each, except small labels in diagrams (under 1.2 in wide), which keep the size the library gave them while text beside them may grow.

## 5. Check the result

- If the slide-review skill is installed, run its script on the new file (`python3 ../slide-review/scripts/review_deck.py "<new file>"`) and pass on anything new, such as a title number that is no longer on the page.
- If you can render pages (LibreOffice: `soffice --headless --convert-to pdf`), look at each rebuilt page. Check that the text fits its boxes and that nothing overlaps.
- Read the plan back against the original slides: every item of each slide is on its new page.

## 6. Report

Lead with the file name of the new deck. Then give one line per slide:

- **Slide 3, rebuilt as Executive summary:** five topics moved into rows; body text down to 11 pt to fit.
- **Slide 5, left as it is:** it has a chart.
- **Slide 6, left as it is:** no page fits seven unrelated points; consider splitting it.

End with anything the user should check by eye, and any suggested message titles.

## Rules

- Change layout, never meaning: no new facts, numbers or claims, and no dropped items.
- The deck's own theme decides fonts and colours. Do not add colours or fonts.
- One run, one new file. Never overwrite the user's deck.
- The scripts read the deck you pass and write only the new file and the report. Nothing is uploaded.

## The pages

`assets/pages.json` holds a sample of the SlideShortcuts slide library: text-led consulting pages (summaries, matrices, processes, columns, lists, plans). The full library of more than 500 pages comes with a SlideShortcuts subscription.
