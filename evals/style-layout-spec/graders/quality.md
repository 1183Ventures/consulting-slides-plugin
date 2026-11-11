---
type: llm
weight: 2
---

The user wants a layout spec, in inches, for a 13.33 x 7.5 in consulting page with a bar chart of 2025 revenue by region and three takeaways.

PASS only if all of these are true:
1. The title sits at the top left, about 0.5 to 0.7 in from the left edge, and is an action title (a full-sentence takeaway), or the response says it should be.
2. The three takeaways go in a panel on the right side of the page (starting somewhere between about 8.5 and 10 in from the left), not floating over or under the chart, and the title is kept narrower than the full width so it does not run across that panel.
3. The chart has an exhibit title or label that states the unit and the period (for example "$M, 2025").
4. There is a source line at the bottom left.
5. It does not put a heavy frame, a 3-D effect or gridlines on the chart; it prefers direct labels or a light baseline.

FAIL if any of these is false.
