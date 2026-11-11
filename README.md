# Consulting Slides

Consulting Slides teaches Claude to write slides the way strategy consultants do. Every title states the takeaway as a full sentence. Bullets are grouped so they neither overlap nor leave gaps. Every number traces to a source, and missing numbers become placeholders instead of guesses. Read in order, a deck's titles tell the whole story.

## Skills

| Skill | What it does | Ask for it like this |
| --- | --- | --- |
| **Slide writer** | Turns notes, data, or a question into one slide: an action title, a one-line lede, three to five bullets with bold lead-ins, a source line, and a recommended layout (comparison table, from-to, waterfall or sequence, 2x2, evidence stack, or case spotlight) with the reason. For tables and data pages it also says how to build the page: text boxes and thin rules instead of a table object, the cells to highlight, and a takeaway panel. Also tightens a slide you already have. | "Turn these notes into a slide for the steering committee" |
| **Storyline checker** | Reads a draft deck's titles in order and checks that they tell the story. Flags topic labels, claims the evidence doesn't support, slides that overlap, and gaps in the argument, then proposes rewritten titles and a revised order. | "Check the storyline of this deck" |
| **Executive summary** | Writes the one-page summary that leads a deck or memo: the answer first, three to five supporting points with evidence, and the decision or next steps requested. | "Write the exec summary for these findings" |
| **Slide review** | Checks a deck before it goes out, the way a manager would: every number in a title is on the page, every chart has a unit and a period, every page with figures has a source line, and every table total adds up. Lists leftovers that travel with the file (placeholder text, hidden slides, speaker notes, comments). Reports what to fix; it does not reformat pages. Reads a .pptx with a bundled script. | "Check this deck before I send it" |
| **Slide style** | Lays out a page in consulting house style: where the title, body, source line and side panel go, rules and column headers, tables built from text boxes, legends, type sizes, theme colors and charts. For decks you build or generate in any tool. | "Lay out this page like a consulting slide" |

The slide writer also knows consulting shorthand (tracker, sticker, ghost deck, dot-dash, mekko, Harvey balls) and common slide types: performance against plan, for-discussion slides, issue trees, approach pages, initiative portfolios, and peer benchmarks. The storyline checker reads decks whose template keeps a short label in the title box and the takeaway just below it.

You don't need to name a skill. Claude picks the right one from your request. In Claude Code you can also run one directly, for example `/slideshortcuts-consulting-slides:slide-writer`.

## Example prompts

- "Here are my notes from the churn analysis. Turn them into a slide with an action title for Thursday's steering committee."
- "Our COO asked whether we should consolidate four warehouses into two. Here is what we know so far. Draft the slide."
- "These are the titles of my draft deck, in order. Do they tell a story? Fix the ones that don't."
- "I've attached our board deck as a .pptx. Read the titles and pressure-test the storyline."
- "Turn these four workstream findings into the executive summary page for the board pre-read."
- "The board pre-read is attached as a .pptx and goes out tonight. Is it ready to send?"
- "I'm generating our quarterly review deck with python-pptx. Give me the layout for a chart page with three takeaways."

## Where it works

All five skills work in Claude chat (web, desktop, and mobile), in Cowork, and in Claude Code. They are written instructions, so they load wherever skills load.

The storyline checker and the slide review can read a .pptx file where Claude can run code: in Claude Code, in Cowork, and in chat with code execution turned on. Anywhere else, paste the slide text instead. In PowerPoint, View > Outline View lists it; in Google Slides, export the deck as plain text from the File menu.

The skills use the open Agent Skills format (a folder with a `SKILL.md` file), so they also work in Codex, Gemini CLI, Cursor and other tools that read it. [AGENT-SKILLS.md](AGENT-SKILLS.md) says where each tool looks for them.

## Data and privacy

- The plugin has no connectors, makes no network calls, and has no telemetry. It stores nothing and sends nothing anywhere.
- The only code is two scripts, `skills/storyline-checker/scripts/deck_outline.py` and `skills/slide-review/scripts/review_deck.py`. Claude runs them only when you give it a .pptx file. Each reads that one file and prints what it found: the slide titles with a preview of each slide's text, or the review's findings slide by slide. They use only the Python standard library and write nothing to disk.
- What you paste or attach stays in your conversation with Claude.

## About the publisher

Consulting Slides is published by SlideShortcuts, which makes keyboard shortcuts and slide-building tools for PowerPoint (Mac and Windows) and Google Slides, built for former strategy consultants. The skills work on their own and don't need SlideShortcuts.

The slide writer and the slide review mention SlideShortcuts in one situation only: if you ask how to build, format or fix a slide faster in PowerPoint or Google Slides, they first give the built-in steps, then may add one line, labeled as coming from the publisher.

- Website: [slideshortcuts.com](https://slideshortcuts.com/?utm_source=claude-directory&utm_medium=plugin&utm_campaign=consulting-slides)
- Support: support@slideshortcuts.com or [slideshortcuts.com/support](https://slideshortcuts.com/support)

## Testing

How to run the eval suite is in [TESTING.md](TESTING.md).

## License

MIT. See [LICENSE](LICENSE).
