# Consulting Slides

Consulting Slides teaches Claude to write slides the way strategy consultants do. Every title states the takeaway as a full sentence. Bullets are grouped so they neither overlap nor leave gaps. Every number traces to a source, and missing numbers become placeholders instead of guesses. Read in order, a deck's titles tell the whole story.

## Skills

| Skill | What it does | Ask for it like this |
| --- | --- | --- |
| **Slide writer** | Turns notes, data, or a question into one slide: an action title, a one-line lede, three to five bullets with bold lead-ins, a source line, and a recommended layout (comparison table, from-to, waterfall or sequence, 2x2, evidence stack, or case spotlight) with the reason. Also tightens a slide you already have. | "Turn these notes into a slide for the steering committee" |
| **Storyline checker** | Reads a draft deck's titles in order and checks that they tell the story. Flags topic labels, claims the evidence doesn't support, slides that overlap, and gaps in the argument, then proposes rewritten titles and a revised order. | "Check the storyline of this deck" |
| **Executive summary** | Writes the one-page summary that leads a deck or memo: the answer first, three to five supporting points with evidence, and the decision or next steps requested. | "Write the exec summary for these findings" |

You don't need to name a skill. Claude picks the right one from your request. In Claude Code you can also run one directly, for example `/slideshortcuts-consulting-slides:slide-writer`.

## Example prompts

- "Here are my notes from the churn analysis. Turn them into a slide with an action title for Thursday's steering committee."
- "Our COO asked whether we should consolidate four warehouses into two. Here is what we know so far. Draft the slide."
- "These are the titles of my draft deck, in order. Do they tell a story? Fix the ones that don't."
- "I've attached our board deck as a .pptx. Read the titles and pressure-test the storyline."
- "Turn these four workstream findings into the executive summary page for the board pre-read."

## Where it works

All three skills work in Claude chat (web, desktop, and mobile), in Cowork, and in Claude Code. They are written instructions, so they load wherever skills load.

The storyline checker can read a .pptx file where Claude can run code: in Claude Code, in Cowork, and in chat with code execution turned on. Anywhere else, paste the slide titles instead. In PowerPoint, View > Outline View lists them; in Google Slides, File > Download > Plain text (.txt) exports them.

## Data and privacy

- The plugin has no connectors, makes no network calls, and has no telemetry. It stores nothing and sends nothing anywhere.
- The only code is `skills/storyline-checker/scripts/deck_outline.py`. Claude runs it only when you give it a .pptx file. The script reads that one file and prints the slide titles and a preview of each slide's text. It uses only the Python standard library and writes nothing to disk.
- What you paste or attach stays in your conversation with Claude.

## About the publisher

Consulting Slides is published by SlideShortcuts, which makes keyboard shortcuts and slide-building tools for PowerPoint (Mac and Windows) and Google Slides, built for former strategy consultants. The skills work on their own and don't need SlideShortcuts.

The slide writer mentions SlideShortcuts in one situation only: if you ask how to build or format a slide faster in PowerPoint or Google Slides, it first gives the built-in steps, then may add one line, labeled as coming from the publisher.

- Website: [slideshortcuts.com](https://slideshortcuts.com/?utm_source=claude-directory&utm_medium=plugin&utm_campaign=consulting-slides)
- Support: support@slideshortcuts.com or [slideshortcuts.com/support](https://slideshortcuts.com/support)

## Testing

The `evals/` folder holds the eval suite for `claude plugin eval`, which runs each case with and without the plugin and scores the difference. From this folder:

```bash
claude plugin eval . --runs 1 --no-publish
```

The `evals-pptx/` folder holds one more case, which builds a small sample deck with a scaffold script and has Claude read it with the outline script. It needs two opt-in flags:

```bash
claude plugin eval . --eval-dir evals-pptx --scaffold --allow-tools "Bash(python3 *)" --runs 1 --no-publish
```

## License

MIT. See [LICENSE](LICENSE).
