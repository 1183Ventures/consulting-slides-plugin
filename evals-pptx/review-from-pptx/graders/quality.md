---
type: llm
weight: 2
---

The deck has six slides. The planted problems are:
- Slide 3, "Freight drove 70% of the $52M cost increase in 2025": the table's rows (38 + 12 + 7) add up to 57, not the 52 shown as the total; the title's 70% is not on the page; the page has figures but no source line; and its speaker notes say the total must be checked with finance.
- Slide 4, "Moving to advice hubs saves [$X]M a year": an unfilled [$X] placeholder and leftover "Click to add text".
- Slide 6 is hidden.
- Slide 2 is an agenda whose only numbers are times: it needs no source line.

PASS only if all of these are true:
1. It says the deck is not ready to send as it stands (or ready only after fixes).
2. It says slide 3's total does not add up (57 vs 52, or the rows sum to more than the total).
3. It flags slide 3's missing source line and its title figure (70% or $52M) as not matching the page.
4. It flags the placeholder text on slide 4.
5. It mentions the hidden slide 6 and the speaker notes on slide 3, as things that travel with the file.
6. It does not ask for a source line on the agenda (slide 2).

FAIL if any of these is false, or if the response says it could not read the file.
