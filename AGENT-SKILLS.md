# Use these skills in any AI agent

The five skills in `skills/` are plain Agent Skills: each is a folder with a `SKILL.md` file, plus any scripts and reference files it needs. Claude Code, Codex, Gemini CLI and Cursor all read this format, and so does any other tool that supports `SKILL.md`.

| Skill | What it does |
| --- | --- |
| `slide-writer` | Turns notes, data or a question into one consulting slide: action title, lede, bullets with bold lead-ins, source line, layout |
| `storyline-checker` | Reads a deck's titles in order and checks that they tell the story |
| `executive-summary` | Writes the one-page summary that leads a deck or memo |
| `slide-review` | Checks a deck before it goes out: title numbers on the page, units and periods on charts, source lines, table totals, leftovers that travel with the file |
| `slide-style` | Lays out a page in consulting house style: title, side panel, rules, tables, legends, type sizes, colors, charts |

## Get the skills

Clone the repository, or download it as a ZIP from GitHub (Code > Download ZIP) and unzip it:

```
git clone https://github.com/1183Ventures/consulting-slides-plugin.git
```

Each tool below needs the folders inside `skills/` copied to the folder it reads. Copy the whole folder for each skill, not just `SKILL.md`: `slide-review` and `storyline-checker` include a script, and `slide-writer` includes reference files.

## Claude Code

Install the Consulting Slides plugin, or copy the skills to your personal skills folder:

```
mkdir -p ~/.claude/skills
cp -R consulting-slides-plugin/skills/* ~/.claude/skills/
```

For one project only, copy them to `.claude/skills/` in that project. Claude picks a skill from your request, or you can type `/slide-review` (or any skill's name) to run it.

## Codex

Copy the skills to your personal skills folder:

```
mkdir -p ~/.agents/skills
cp -R consulting-slides-plugin/skills/* ~/.agents/skills/
```

For one repository only, copy them to `.agents/skills/` at its root. Restart Codex if they don't appear. Codex picks a skill from your request, or you can name one, for example `$slide-review`.

## Gemini CLI

Install all five from the repository in one command:

```
gemini skills install https://github.com/1183Ventures/consulting-slides-plugin.git --path skills
```

Add `--scope workspace` to install them in the current project only. Or copy the folders to `~/.gemini/skills/` (or `~/.agents/skills/`) yourself. Check with `/skills list`.

## Cursor

Copy the skills to `~/.cursor/skills/` or `~/.agents/skills/` for every project, or to `.cursor/skills/` or `.agents/skills/` in one project. Cursor also reads `~/.claude/skills/`, so a copy made for Claude Code works there too. Type `/` in Agent chat and search for the skill's name to run it directly.

## One copy for every tool

Codex, Gemini CLI and Cursor all read `~/.agents/skills/`, and Claude Code and Cursor both read `~/.claude/skills/`. Two copies cover all four tools:

```
mkdir -p ~/.agents/skills ~/.claude/skills
cp -R consulting-slides-plugin/skills/* ~/.agents/skills/
cp -R consulting-slides-plugin/skills/* ~/.claude/skills/
```

To update, pull the repository again (`git pull`) and copy the folders again.

## Scripts

Two skills include a short Python script. Both need Python 3.8 or later, use only the standard library, read only the .pptx file you give them, write nothing to disk, and make no network calls. Your agent runs them only when you give it a .pptx file, and asks first if your settings require it.

- `storyline-checker/scripts/deck_outline.py` prints each slide's title in order, with a preview of its text.
- `slide-review/scripts/review_deck.py` prints, slide by slide, a title number that is not on the page, a chart without a unit or period, a page with figures and no source line, a table total that doesn't add up, and leftovers such as placeholder text, hidden slides, speaker notes and comments.

Without code execution, paste the slide text instead, or share page images for the slide review's layout checks.

## Optional files

`agents/openai.yaml` inside each skill holds a display name and a short description for Codex and the ChatGPT desktop app. Other tools ignore it.
