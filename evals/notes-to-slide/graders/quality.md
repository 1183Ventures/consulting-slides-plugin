---
type: llm
weight: 2
---

The user asked for one slide built from churn notes. The notes contain only these numbers: 3.2%, 61%, 90 days, 1.1%, 4.8%, 44%, 18%, 52%, 0.2 points, 1,240, 2025, Q3.

PASS only if all of these are true:
1. The slide title is one full sentence that states a conclusion (it has a verb and a so-what, for example that onboarding rather than price is the lever) and includes at least one number from the notes. A topic label such as "Churn analysis" or "Churn deep dive" fails.
2. There is a short lede of one or two sentences between the title and the bullets.
3. The body has three to five bullets, and no two bullets make the same point.
4. Every number in the response is from the list above, follows from it by simple arithmetic (such as 4.8% against 1.1% being about four times, or 56% not completing onboarding), or is clearly marked as a placeholder or estimate.
5. A source line mentions the billing system or the exit survey, or says the source must be confirmed.

FAIL if any of these is false.
