# PPTX Template Editing

A pattern for building branded PowerPoint presentations by having an LLM edit the XML inside a real template file.

This is an idea file. Give it to your LLM agent (Claude Code, Codex, etc.) along with a PowerPoint template. The agent will set everything up, build the reference material with you, and then produce decks on demand.

## The core idea

Most LLM-generated presentations are disappointing. The LLM generates slides from code (python-pptx, pptxgenjs), producing bland output that looks nothing like what your design team made. Your company has a beautiful template with exact colors, fonts, shapes, and layouts — but the LLM ignores all of it.

The idea here is different. Instead of generating slides from code, the LLM **edits the actual PowerPoint template file**. A .pptx is a ZIP of XML files. Unpack it, manipulate the XML to select slides, reorder them, replace placeholder text — repack it. The output preserves every brand element exactly: colors, gradients, shapes, font stacks, image positions, rounded corners, accent bars. The LLM is just swapping the words.

**PowerPoint templates are programs. The slides are functions. The placeholder text is the arguments.** The LLM's job is to call the right functions with the right arguments — not to write the program from scratch.

## Setup

When the user gives you a .pptx template for the first time, run through this setup process. It produces a workspace with reference files that persist across sessions.

### Step 1: Create the workspace

Create this directory structure next to or inside the user's project:

```
pptx-workspace/
  template/
    template.pptx          # the original template (never modified)
  references/
    layout-map.md           # slide inventory — you and the user build this
    brand-guide.md          # colors, fonts, styling — extracted from theme XML
    slide-text-dump.md      # all text from every slide — used during setup
  output/                   # finished decks go here
```

Copy the user's .pptx into `template/template.pptx`.

### Step 2: Extract the theme

Unpack the template and read `ppt/theme/theme1.xml`. Write `references/brand-guide.md` with:

```python
import zipfile, os, re
from xml.etree import ElementTree as ET

# Unpack
with zipfile.ZipFile('pptx-workspace/template/template.pptx', 'r') as z:
    z.extractall('/tmp/pptx-inspect/')

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# Dimensions
tree = ET.parse('/tmp/pptx-inspect/ppt/presentation.xml')
sld_sz = tree.getroot().find('.//p:sldSz', NS)
cx, cy = int(sld_sz.get('cx', '9144000')), int(sld_sz.get('cy', '5143500'))

# Colors
theme = ET.parse('/tmp/pptx-inspect/ppt/theme/theme1.xml')
clr_scheme = theme.getroot().find('.//a:clrScheme', NS)
colors = {}
for role in ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6']:
    elem = clr_scheme.find('a:{}'.format(role), NS)
    if elem is not None:
        srgb = elem.find('a:srgbClr', NS)
        sys_clr = elem.find('a:sysClr', NS)
        if srgb is not None:
            colors[role] = srgb.get('val')
        elif sys_clr is not None:
            colors[role] = sys_clr.get('lastClr', sys_clr.get('val', ''))

# Fonts
font_scheme = theme.getroot().find('.//a:fontScheme', NS)
major = font_scheme.find('a:majorFont/a:latin', NS).get('typeface', '?')
minor = font_scheme.find('a:minorFont/a:latin', NS).get('typeface', '?')
```

Write the results into `references/brand-guide.md` in this format:

```markdown
# Brand Guide

## Dimensions
[width] x [height] EMU ([inches] x [inches])

## Theme Colors
| Role | Hex | Name | Usage |
|------|-----|------|-------|
| dk1 | #XXXXXX | TODO: name | TODO: usage |
...

## Background Colors
| Hex | Label | Usage |
|-----|-------|-------|
| TODO | Dark | Primary dark background |
| TODO | Light | Primary light background |

## Fonts
- Heading: [major font]
- Body: [minor font]

## Card Styling
TODO: filled in after the user classifies layouts
```

Tell the user: "I've extracted your template's colors and fonts. I'll need your help naming them and filling in the usage rules after we classify the slides."

### Step 3: Dump all text from every slide

This is critical — it's how you and the user classify layouts together.

```python
import re, os

slides_dir = '/tmp/pptx-inspect/ppt/slides/'
dump_lines = ['# Slide Text Dump\n']

slide_files = sorted(
    [f for f in os.listdir(slides_dir) if f.startswith('slide') and f.endswith('.xml')],
    key=lambda f: int(re.search(r'(\d+)', f).group())
)

for fname in slide_files:
    num = re.search(r'(\d+)', fname).group()
    with open(os.path.join(slides_dir, fname), 'r', encoding='utf-8') as f:
        content = f.read()

    # Extract all text
    texts = re.findall(r'<a:t>([^<]+)</a:t>', content)

    # Detect background
    bg_match = re.search(r'<p:bg>.*?<a:srgbClr\s+val="([A-Fa-f0-9]{6})"', content, re.DOTALL)
    bg = '#{}'.format(bg_match.group(1)) if bg_match else 'inherited'

    dump_lines.append('## Slide {}'.format(num))
    dump_lines.append('Background: {}'.format(bg))
    dump_lines.append('')
    for t in texts:
        t = t.strip()
        if t:
            dump_lines.append('- {}'.format(t))
    dump_lines.append('')

with open('pptx-workspace/references/slide-text-dump.md', 'w') as f:
    f.write('\n'.join(dump_lines))
```

Now present the dump to the user slide by slide and ask them to classify each one. Say something like:

> Here's what I found in your template. I need your help classifying each slide.
> For each one, tell me: what type of layout is this?
>
> **Slide 1** (bg: inherited)
> - Presentation Title
> - Subtitle goes here
> - September 25, 2024
>
> This looks like a **Cover** slide. Is that right?
>
> **Slide 2** (bg: #1D1925 — dark)
> - THE OPPORTUNITY
>
> This looks like a **Section Divider**. Is that right?

Work through all slides. Group similar slides together. Common layout types:
- **Cover / Title** — opening slides with title, subtitle, date
- **Section Divider** — dark slide with centered text, used between sections
- **Agenda** — text list, often with an image on one side
- **Content Cards** — 3-4 cards in a row for comparing items
- **Numbered Rows** — sequential items (01, 02, 03) with descriptions
- **Stacked Rows** — framework/hierarchy with labeled tiers
- **Stats / Big Numbers** — large metric callouts with labels
- **Timeline** — horizontal flow with phases or milestones
- **Color Bumper** — gradient background slides for section breaks or closing
- **Text + Image** — narrative on one side, image on the other
- **Quote** — testimonial or inspirational quote
- **People / Team** — headshots with names

For slides that are dark/light pairs of the same layout (e.g., slide 24 is dark cards, slide 34 is the same layout on a light background), note both.

### Step 4: Build the layout map

From the classification session, write `references/layout-map.md`:

```markdown
# Layout Map

## Preset Deck Structures

### Quarterly Review (8 slides)
1. Cover (slide ??, light) — "[Account] QBR [Quarter]"
2. Agenda (slide ??, light) — Meeting outline
3. Stats (slide ??, light) — Key metrics
...

### Project Update (6 slides)
...

## Layout Types

### COVER
- Slide 4 (light bg) — title + subtitle + date
- Slide 3 (dark bg) — title + date only

### SECTION DIVIDER
- Slide 2 (dark bg) — centered pill badge, one line

### AGENDA
- Slide 14 (light bg) — left text + right image

[...continue for all types...]

## Slide Selection Decision Tree
1. Opening? → Cover
2. New section? → Section Divider or Color Bumper
3. Agenda? → Agenda
4. Comparing items? → Content Cards
5. Sequential steps? → Numbered Rows
6. Key metrics? → Stats
7. Timeline? → Timeline
8. Closing? → Color Bumper

## Text Constraints
| Layout | Element | Max Length | Notes |
|--------|---------|-----------|-------|
| Cover | Title | ~8 words | |
| Stats | Big number | 3-4 chars | Overlaps with breadcrumb |
| Cards | Card body | 2-3 sentences | Fixed-width box |
| Monospace labels | Label | ~12 chars | Very narrow |
```

Ask the user to fill in the preset deck structures based on what kinds of decks their team commonly builds.

### Step 5: Sanitize the template

PowerPoint templates contain smart quotes, em dashes, and split text runs that break string matching. Sanitize once and save the clean version as the working template.

```python
import zipfile, re, os, shutil

SMART_CHARS = {
    '\u2018': "'", '\u2019': "'",   # smart single quotes
    '\u201C': '"', '\u201D': '"',   # smart double quotes
    '\u2013': '-', '\u2014': '--',  # en/em dash
    '\u2026': '...', '\u00A0': ' ', # ellipsis, nbsp
}

# Work on a copy
shutil.copytree('/tmp/pptx-inspect/', '/tmp/pptx-sanitized/')
slides_dir = '/tmp/pptx-sanitized/ppt/slides/'

changes = 0
for fname in sorted(os.listdir(slides_dir)):
    if not fname.endswith('.xml'):
        continue
    fpath = os.path.join(slides_dir, fname)
    with open(fpath, 'r', encoding='utf-8') as f:
        content = f.read()
    original = content

    # Replace smart characters inside <a:t> elements
    def replace_smart(match):
        text = match.group(1)
        for smart, plain in SMART_CHARS.items():
            text = text.replace(smart, plain)
        return '<a:t>{}</a:t>'.format(text)
    content = re.sub(r'<a:t>(.*?)</a:t>', replace_smart, content, flags=re.DOTALL)

    # Replace standalone & with 'and' in text nodes (not XML entities)
    def replace_amp(match):
        text = match.group(1)
        text = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#)', 'and', text)
        return '<a:t>{}</a:t>'.format(text)
    content = re.sub(r'<a:t>(.*?)</a:t>', replace_amp, content, flags=re.DOTALL)

    if content != original:
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(content)
        changes += 1

# Repack as sanitized template
with zipfile.ZipFile('pptx-workspace/template/template.pptx', 'w', zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk('/tmp/pptx-sanitized/'):
        for file in files:
            file_path = os.path.join(root, file)
            arcname = os.path.relpath(file_path, '/tmp/pptx-sanitized/')
            z.writestr(arcname, open(file_path, 'rb').read())

print('{} slide files sanitized'.format(changes))
```

Tell the user: "I've sanitized your template — smart quotes, special characters, and ampersands are normalized. The clean version is saved as the working template."

### Step 6: Clean up

```python
import shutil
shutil.rmtree('/tmp/pptx-inspect/', ignore_errors=True)
shutil.rmtree('/tmp/pptx-sanitized/', ignore_errors=True)
```

Tell the user: "Setup is complete. Your workspace is ready. From now on, just tell me what deck you need and I'll build it."

## Producing a deck

Once setup is complete, follow this process every time the user asks for a presentation.

### 1. Plan

Read `references/layout-map.md`. Based on the user's request, propose a slide plan:

```
Slide 1: Cover (slide 4, light) — "Q2 Deployment Review"
Slide 2: Agenda (slide 14, light) — Meeting outline
Slide 3: Stats (slide 33, light) — Key deployment metrics
...
```

Check text constraints from the layout map. If any content exceeds limits, summarize it. Ask the user to confirm the plan before building.

### 2. Build

Unpack the template to a working directory:

```python
import zipfile
with zipfile.ZipFile('pptx-workspace/template/template.pptx', 'r') as z:
    z.extractall('/tmp/pptx-build/')
```

Extract the rId mapping:

```python
import re
with open('/tmp/pptx-build/ppt/_rels/presentation.xml.rels', 'r') as f:
    rels = f.read()
entries = re.findall(r'Id="(rId\d+)"[^>]*Target="(slides/slide\d+\.xml)"', rels)
rid_to_slide = {rid: target for rid, target in entries}
```

Modify `ppt/presentation.xml` to select and reorder slides. **Use regex — never xml.dom.minidom** (it corrupts the XML prolog):

```python
with open('/tmp/pptx-build/ppt/presentation.xml', 'r') as f:
    pres = f.read()

all_entries = re.findall(r'<p:sldId[^/]*/>', pres)
rid_to_entry = {}
for entry in all_entries:
    rid = re.search(r'r:id="(rId\d+)"', entry).group(1)
    rid_to_entry[rid] = entry

keep_rids = ['rId6', 'rId16', 'rId35']  # from your plan
new_entries = [rid_to_entry[r] for r in keep_rids if r in rid_to_entry]
new_block = '<p:sldIdLst>' + ''.join(new_entries) + '</p:sldIdLst>'
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', new_block, pres, flags=re.DOTALL)

with open('/tmp/pptx-build/ppt/presentation.xml', 'w') as f:
    f.write(pres)
```

### 3. Edit content

Replace ALL placeholder text using `safe_replace()`. This function targets only `<a:t>` text nodes — it will not corrupt XML attributes like EMU coordinates or relationship IDs:

```python
def safe_replace(content, old_text, new_text):
    """Replace text only inside <a:t> elements."""
    wrapped_old = '<a:t>{}</a:t>'.format(old_text)
    wrapped_new = '<a:t>{}</a:t>'.format(new_text)
    if wrapped_old in content:
        return content.replace(wrapped_old, wrapped_new), True
    if len(old_text) > 20 and old_text in content:
        return content.replace(old_text, new_text), True
    pattern = re.compile(
        r'(<a:t>[^<]*?)' + re.escape(old_text) + r'([^<]*?</a:t>)')
    if pattern.search(content):
        return pattern.sub(r'\g<1>' + new_text + r'\g<2>', content), True
    return content, False
```

Build ALL replacements, then execute in one pass per slide. **Sort by old_text length descending** — this prevents shorter strings from breaking longer ones:

```python
replacements = {
    'ppt/slides/slide4.xml': [
        ('Presentation title', 'Q2 Deployment Review'),
        ('Subtitle text here', 'Acme Corp — June 2026'),
    ],
    # ... every slide in the plan
}

for slide_file, pairs in replacements.items():
    with open('/tmp/pptx-build/{}'.format(slide_file), 'r') as f:
        content = f.read()
    for old_text, new_text in sorted(pairs, key=lambda p: len(p[0]), reverse=True):
        content, found = safe_replace(content, old_text, new_text)
        if not found:
            print('WARNING: not found: {}'.format(old_text[:50]))
    with open('/tmp/pptx-build/{}'.format(slide_file), 'w') as f:
        f.write(content)
```

**Replace EVERY text element** — titles, breadcrumbs, body text, labels, stat numbers, footnotes. Anything left untouched shows the original template placeholder. Use `references/slide-text-dump.md` to know exactly what text exists on each slide.

### 4. Pack and deliver

Repack directly with `zipfile`. Do not use any other packing tool — some strip the XML prolog and corrupt the file:

```python
import zipfile, os
output_name = 'q2-deployment-review.pptx'  # descriptive name
output_path = 'pptx-workspace/output/{}'.format(output_name)
with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk('/tmp/pptx-build/'):
        for file in files:
            file_path = os.path.join(root, file)
            arcname = os.path.relpath(file_path, '/tmp/pptx-build/')
            z.writestr(arcname, open(file_path, 'rb').read())
```

### 5. QA

Extract text and verify:

```python
# Quick text extraction for QA
import re, os
for fname in sorted(os.listdir('/tmp/pptx-build/ppt/slides/')):
    if not fname.endswith('.xml'):
        continue
    with open('/tmp/pptx-build/ppt/slides/{}'.format(fname), 'r') as f:
        content = f.read()
    texts = [t.strip() for t in re.findall(r'<a:t>([^<]+)</a:t>', content) if t.strip()]
    if texts:
        print('--- {} ---'.format(fname))
        for t in texts:
            print('  {}'.format(t))
```

Check:
- No leftover template placeholder text
- Slide count matches the plan
- Breadcrumbs updated on every slide
- `presentation.xml` still has `standalone="yes"` in the prolog

If anything is structurally wrong, **don't patch — delete `/tmp/pptx-build/` and start from step 2.** Template-based editing is fast; debugging corrupted PPTX XML is slow.

### 6. Clean up

```python
import shutil
shutil.rmtree('/tmp/pptx-build/', ignore_errors=True)
```

Tell the user where the file is and that it's ready to open.

## Pitfalls reference

These are the hard lessons that took multiple iterations to learn. They are the reason this document exists. Follow them exactly.

**Never use xml.dom.minidom on presentation.xml.** It strips `standalone="yes"` from the XML prolog. PowerPoint checks for this and triggers a repair dialog that can silently drop slides. Use regex or plain string operations.

**Never use plain str.replace() for short strings.** `content.replace('12', '95')` corrupts the file. PPTX XML has thousands of numeric values in attributes (`<a:off x="1200150"/>`). Always use `safe_replace()` which targets only `<a:t>` text nodes.

**Sort replacements longest-first.** If "Employee Success" and "Success" are both targets, replacing "Success" first breaks the longer match.

**Replace everything or it leaks through.** Every text element on every selected slide must be replaced. Forgetting a breadcrumb means the output says "Marketing Team Quarterly Priorities" on a slide about deployment metrics.

**Sanitize the template once.** Smart quotes (`\u201C`), em dashes (`\u2014`), and split text runs (`<a:t>Hello </a:t><a:t>World</a:t>`) make str.replace() fail silently. The setup process handles this.

**Error recovery: nuke and re-unpack.** If anything goes wrong structurally, delete the working directory and unpack fresh from the template. Never try to patch corrupted XML.

## Tips for good decks

- **Dark/light rhythm**: alternate backgrounds for visual variety. Avoid 3+ consecutive same-background slides.
- **Stats slides are tricky**: big numbers often overlap breadcrumbs by design. Keep numbers to 3-4 characters, titles to 2 words.
- **Monospace labels are narrow**: single words work best for uppercase monospace tags.
- **Cover images get stripped**: use section dividers or color bumpers instead of image-dependent covers.
- **Content must be scannable**: card content is 2-3 short sentences max. Summarize — never overflow.

## Note

This document is intentionally focused on the technique rather than a specific template. The exact slide numbers, layout types, and colors will be different for every template. Share this file with your LLM agent along with your PowerPoint template — the agent will run through the setup, build the reference material with you, and then produce decks on demand from that point forward. The document's job is to give the agent the right process and transfer the hard-won pitfalls. The agent can figure out the rest.
