# pptx-skill

A pattern for building branded PowerPoint presentations with LLMs — by editing the real template XML instead of generating slides from code.

## The problem

LLM-generated presentations are ugly. Tools like python-pptx and pptxgenjs produce bland, unbranded slides that look nothing like what your design team made. Your company's beautiful template — exact colors, fonts, shapes, layouts — gets ignored entirely.

## The idea

A .pptx file is just a ZIP of XML files. Instead of generating slides from code, have the LLM **edit your actual template**: unpack it, select and reorder slides, swap the placeholder text, repack it. The output preserves every brand element exactly — gradients, rounded corners, accent bars, font stacks — because the LLM is editing the design, not recreating it.

**PowerPoint templates are programs. Slides are functions. Placeholder text is arguments.**

## How it works

1. **Give your LLM agent [`PPTX-TEMPLATE-EDITING.md`](PPTX-TEMPLATE-EDITING.md) and your .pptx template**
2. The agent runs a one-time **setup**: extracts your theme colors and fonts, dumps all text from every slide, and walks you through classifying each slide's layout type (Cover, Agenda, Stats, Timeline, etc.)
3. From then on, just ask for a deck. The agent **plans** the slides, **unpacks** the template, **edits** the XML, **repacks** it, and **QA checks** the output — all automatically

The guide includes working Python code for every step and documents the hard-won pitfalls that took multiple iterations to discover (XML prolog corruption, short-string replacement bombs, smart quote failures, and more).

## Quick start

```bash
# Clone this repo
git clone https://github.com/mauriciomcv/pptx-skill.git

# Open Claude Code (or your LLM agent of choice) in the repo
claude

# Then say:
# "Read PPTX-TEMPLATE-EDITING.md and set up my template"
# (drag and drop or provide the path to your .pptx file)
```

The agent will create a `pptx-workspace/` directory with your sanitized template and reference files. After setup, just ask for decks:

> "Create a 6-slide deck for the Q2 deployment review with key metrics, timeline, and next steps."

## What's in the guide

| Section | What it covers |
|---------|---------------|
| **Setup** (one-time) | Create workspace, extract theme, dump slide text, classify layouts with the user, sanitize template |
| **Producing a deck** (repeatable) | Plan slides, unpack, select/reorder, batch-edit text, pack, QA |
| **Pitfalls** | minidom corruption, safe_replace(), sort-by-length, replace-everything, sanitization, error recovery |
| **Tips** | Dark/light rhythm, stats overlaps, monospace constraints, content sizing |

## What you need

- An LLM agent that can run Python (Claude Code, Codex, OpenCode, etc.)
- A PowerPoint template (.pptx) from your organization
- ~15 minutes for the one-time setup (classifying your slides)

## What you get

On-brand presentations that preserve your template's exact design — produced in seconds instead of hours.

## Tested with

- Claude Code (Claude Opus / Sonnet)
- Templates ranging from 20 to 60+ slides
- Corporate decks, QBRs, project updates, pitch decks

## License

MIT

---

*Inspired by the [LLM Wiki](https://github.com/tobi/llm-wiki) pattern — the idea that LLMs are best used to build and maintain persistent artifacts, not just answer questions.*
