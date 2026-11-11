---
type: llm
weight: 2
---

The user pasted three pages and asked for a check before sending. The planted problems are:
- Page 5: the table's rows (38 + 12 + 7) add up to 57, not the 52 shown as the total.
- Page 5: freight is 38 of the increase, which is 73% of 52 or 67% of 57, not the 70% the title claims; and the title's $52M depends on the wrong total.
- Page 6: the chart has a unit (%) but no period (no year, quarter or date), and the page has figures but no source line.
- Page 2 is an agenda: its numbers are times, so it needs no source line.

PASS only if all of these are true:
1. It says page 5's total does not add up, and gives 57 (or says the rows sum to more than 52).
2. It flags that page 5's title figures (70% or $52M) do not match the table, without silently picking the right number: it asks, marks the figure to check, or shows the arithmetic.
3. It says page 6's chart needs a period, and that page 6 needs a source line.
4. It does not ask for a source line on page 2, the agenda.
5. It does not invent a source, a period or a corrected figure as fact; any it suggests is clearly a placeholder or an option to confirm.

FAIL if any of these is false.
