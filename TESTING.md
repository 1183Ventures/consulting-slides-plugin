# Testing

The `evals/` folder holds the eval suite for Claude Code's plugin evaluator, which runs each case with and without the plugin and scores the difference. From this folder, run `claude plugin eval . --runs 1 --no-publish`.

The `evals-pptx/` folder holds one more case: it builds a small sample deck with a scaffold script and has Claude read it with the outline script. It needs two opt-in flags: `claude plugin eval . --eval-dir evals-pptx --scaffold --allow-tools "Bash(python3 *)" --runs 1 --no-publish`.
