---
name: storyline-checker
description: Checks the storyline of a draft deck by reading its slide titles in order and testing whether the titles alone tell the story. Flags topic-label titles, titles that claim more than the evidence supports, slides that overlap, and gaps in the argument (MECE), then proposes rewritten titles and a revised order. Reads titles pasted as text or an outline, or a .pptx file when Claude can run code. Use when the user asks to review, pressure-test, or tighten a deck; to check its flow, storyline, or horizontal logic; to see whether the titles tell a story; or to review a ghost deck, storyboard, or outline.
---

# Storyline checker

A finished deck can be understood from its titles alone. Read the titles the way a busy partner or client would, find where the story breaks, and propose the fix.

## 1. Get the titles in order

- **Pasted text or an outline:** take one title per line, in the order given, and keep any slide numbers. If body text sits under a title, keep it with that slide; you will use it to test the claims.
- **A .pptx file:** if you can run code, extract the outline with the bundled script:

  ```
  python3 ${CLAUDE_SKILL_DIR}/scripts/deck_outline.py "<path to the .pptx>"
  ```

  In claude.ai chat the script is at `scripts/deck_outline.py` next to this file, so run it from there. It prints each slide's number and title in order with a short preview of the slide's other text, and marks hidden slides and slides whose title had to be guessed. It uses only the Python standard library and reads nothing but the file you pass.
- **No way to read the file:** ask the user to paste the titles. In PowerPoint, View > Outline View lists them in order. In Google Slides, File > Download > Plain text (.txt) exports the deck's text.
- Set aside the cover, agenda, section dividers, and appendix: they are labels by design. Check them only for consistency, for example whether the agenda matches the sections that follow. Leave hidden slides out of the story and mention them once.

## 2. Read the titles alone

Before judging any single title, read only the titles from top to bottom and write the story they tell in three to five sentences. If you can't, that is the main finding: say so plainly.

Then name the governing thought: the one answer the deck exists to give. Say whether a title states it, and where. The answer belongs early, in the executive summary or the first content slide, not on the last page.

## 3. Test each title

Flag a title when it is:

- **A topic label:** a noun phrase with no claim, such as "Market overview" or "Survey results". Every content slide needs a takeaway.
- **Vague:** it has a verb but no so-what ("Costs have changed over time").
- **Unsupported:** it claims more than the slide shows: cause from correlation, a forecast stated as fact, a general rule from one case, or a number that is nowhere in the deck. If you have the slide's body text, check the title against it. If you only have titles, list the claim under "Claims to check" with the evidence it needs.
- **Overloaded:** two or more ideas joined by "and", or longer than about 20 words.
- **Off-style:** a question where the rest are statements, or a different tense, number format, or term from the rest of the deck.

## 4. Test the set

- **Flow:** each title should follow from the one before it. "So", "because", or "but" should fit between them. Flag the jumps.
- **Overlap:** two slides that make the same point or cover the same ground. Propose a merge, or sharpen each so they differ.
- **Gaps:** steps the argument needs that no slide covers. List what a skeptical decision maker must believe to accept the governing thought (for example: the problem is big enough, the cause is known, the fix works, it pays back, it can be done, the risks are manageable, here is the ask) and check that each has a slide.
- **Grouping:** sections should split on one dimension, and no slide should fit in two sections.
- **Order:** the answer first, then the supporting arguments in a deliberate order (most important first, or the order the logic runs), then the recommendation and the ask.
- **Ending:** the deck should end on a decision or next steps with owners and timing, not on a recap.

## 5. Propose rewrites

- Rewrite every flagged title as a full-sentence takeaway: subject, verb, and so-what, specific where the material allows.
- Use only facts and numbers from the user's material. Where a strong title needs a number you don't have, write a placeholder in square brackets, such as [X%], and add it to "Claims to check". Never invent evidence to make a title sound stronger.
- For a gap, write the title of the missing slide and mark it [New].
- Keep the user's terms and the deck's style.
- Never use em dashes or en dashes (Unicode U+2014 and U+2013) as punctuation, in the rewrites or in your review. Use a comma, a colon, parentheses, or two sentences.

## Output format

**Storyline as the titles tell it:** three to five sentences, or a plain statement that the titles don't yet tell a story, and why.

**Governing thought:** where the deck states it, or that it is missing, plus your proposed version.

**Title by title:**

| # | Current title | Issue | Proposed title |
| --- | --- | --- | --- |

List every content slide for decks of up to 25 slides, with "Keep" in the Issue column for titles that pass. For longer decks, list only the flagged ones.

**Deck-level issues:** bullets with bold lead-ins (**Overlap:**, **Gap:**, **Order:**, **Ending:**), each naming the slides involved.

**Revised storyline:** the proposed titles in the proposed order, numbered, with new slides marked [New] and merged slides marked, for example, [Merged from 6 and 7].

**Claims to check:** the titles whose evidence must be confirmed before the deck goes out, each with the evidence it needs.

## Example

Draft titles: 1. Branch network overview. 2. Branch visits fell 31% since 2022 while digital logins doubled. 3. Customer research. 4. Customers still want branches for mortgages and small business accounts. 5. Recommendations.

Review, abridged:
- **Slide 1, topic label:** proposed "Our branches cost [$X]M a year to serve a shrinking share of transactions".
- **Slides 3 and 4, overlap:** slide 3 is a label for the research that slide 4 already states. Merge them: "Customers still want a branch for mortgages and small business accounts, not for everyday banking".
- **Gap:** nothing shows which branches to change or what that saves. [New] "Moving to [N] advice hubs saves [$X]M a year and keeps [Y]% of customers within 15 minutes of a branch".
- **Slide 5, topic label and no ask:** proposed "We recommend piloting advice hubs in two regions, with a go or no-go decision in [quarter]".
