---
type: llm
weight: 2
---

The deck's slide titles, in order, are: 1 "Project Atlas: steering committee update", 2 "Agenda", 3 "Background", 4 "Operating costs rose 12% in 2025, and freight drove two thirds of the increase", 5 "Freight analysis", 6 (hidden) "Old appendix: 2023 cost baseline", 7 "Next steps". The agenda promises "What we recommend", but no slide makes a recommendation.

PASS only if all of these are true:
1. The response works from the deck's real titles: it quotes or clearly refers to "Freight analysis" and to the slide about operating costs rising 12%.
2. It flags "Background" and "Freight analysis" as titles without a takeaway and proposes full-sentence rewrites for them.
3. It points out that no slide makes a recommendation, that the deck ends without a clear ask, or that the agenda's "What we recommend" has no matching slide.

FAIL if any of these is false, or if the response says it could not read the file.
