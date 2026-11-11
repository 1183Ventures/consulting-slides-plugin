# Testing

The `evals/` folder holds the eval suite for Claude Code's plugin evaluator, which runs each case with and without the plugin and scores the difference. From this folder, run `claude plugin eval . --runs 1 --no-publish`.

The `evals-pptx/` folder holds two more cases: each builds a small sample deck with a scaffold script, and has Claude read it with the outline script (storyline checker) or the review script (slide review). They need two opt-in flags: `claude plugin eval . --eval-dir evals-pptx --scaffold --allow-tools "Bash(python3 *)" --runs 1 --no-publish`.

The slide review's script follows the rules of SlideShortcuts' Check Pages. Its page judge (`judge_page` in `skills/slide-review/scripts/review_deck.py`) was checked against that tool's own code on 10,000 random pages, with no differences, on 11 Nov 2026, after Check Pages stopped counting clock times and durations as figures for the source-line check.
