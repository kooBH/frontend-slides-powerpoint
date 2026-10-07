# PowerPoint Slides

A coding-agent skill for creating distinctive, fully editable **PowerPoint (.pptx)** decks — from scratch or by restyling an existing PowerPoint file. It is packaged as a Claude Code plugin, and the core `SKILL.md` can also be read by other coding agents with filesystem and shell access.

This is a PowerPoint-only fork of [zarazhangrui/frontend-slides](https://github.com/zarazhangrui/frontend-slides). The original produces animated HTML decks; this version keeps its workflow and design philosophy ("show, don't tell", anti-AI-slop, progressive disclosure, the bold template pack) and replaces the HTML output with native PowerPoint: real theme colors, slide layouts and placeholders, speaker notes, native charts and tables, transitions, and entrance animations.

## What This Does

**PowerPoint Slides** helps non-designers create beautiful decks without fighting PowerPoint. It uses a "show, don't tell" approach: instead of asking you to describe your aesthetic preferences in words, it builds three candidate title slides, renders them to images, and lets you pick what you like. Then it generates the whole deck in that style.

### Key Features

- **Native .pptx output** — Opens in PowerPoint, Keynote, Google Slides, and LibreOffice. Everything is a real, editable object.
- **Structured, restylable decks** — Theme colors and fonts are written into the file, and every slide frame is a named layout with placeholders, so Design → Variants and View → Slide Master work the way they should.
- **Visual Style Discovery** — Can't articulate design preferences? Pick from rendered previews.
- **Restyle an existing PowerPoint** — Extract every slide's text, images, tables, chart data, and notes, then rebuild in a new design.
- **Anti-AI-Slop** — Curated distinctive styles that avoid generic aesthetics (no Calibri-on-white, no blue title bars, no accent lines under titles).
- **Bold Template Pack** — 34 optional design-forward systems from `beautiful-html-templates`, loaded progressively so safe presets still work as the default fallback.
- **Motion that stays editable** — Consistent transitions and one orchestrated, staggered entrance per slide, injected as native PowerPoint animations you can change in the Animations pane.
- **Built-in QA** — Structural validation, an outline dump, and pixel-accurate rendering (PowerPoint itself on Windows, LibreOffice elsewhere) so the agent looks at every slide before handing it over.
- **Regenerable** — Each deck ships with the `build-deck.js` that produced it. Change a color or a sentence and rebuild.

## Requirements

- A local coding agent with filesystem access and the ability to run shell commands (Claude Code is required only for the marketplace install and the `/powerpoint-slides:powerpoint-slides` command)
- **Node.js 18+** with `pptxgenjs` (`npm install pptxgenjs` in the deck folder; the skill does this for you)
- **Python 3.9+** with `python-pptx` (restyling existing decks) and `Pillow` (background images, contact sheets): `pip install python-pptx Pillow`
- **For rendering previews and QA:** Microsoft PowerPoint on Windows (used through COM automation) **or** [LibreOffice](https://www.libreoffice.org/download/) on any platform. Per-slide PNGs from LibreOffice additionally need `pip install pymupdf` or Poppler's `pdftoppm`.
- Fonts: the distinctive fonts each style names must be installed to render as designed. The skill checks with `scripts/list-fonts.py`, tells you what is missing, and uses the style's Office-safe fallback pair meanwhile.

## Installation

### Via Claude Code Custom Marketplace Source

Install directly from this GitHub repo. Run these as two separate Claude Code messages; do not paste both lines into the prompt at once.

```text
/plugin marketplace add https://github.com/KooBH/frontend-slides-powerpoint
```

After that finishes, run:

```text
/plugin install powerpoint-slides@powerpoint-slides
```

Use the HTTPS URL. The shorter `KooBH/frontend-slides-powerpoint` form may make Claude Code try SSH, which can fail if GitHub is not already in your `known_hosts` file.

Then use it by typing `/powerpoint-slides:powerpoint-slides` in Claude Code. Claude Code namespaces plugin-installed skills as `/plugin-name:skill-name`.

### Claude Code Manual Installation

Copy the skill files to your Claude Code skills directory:

```bash
# Create the skill directory
mkdir -p ~/.claude/skills/powerpoint-slides/scripts

# Copy the user-facing skill files
cp SKILL.md STYLE_PRESETS.md pptx-template.md animation-patterns.md ~/.claude/skills/powerpoint-slides/
cp -R bold-template-pack ~/.claude/skills/powerpoint-slides/
cp scripts/*.py ~/.claude/skills/powerpoint-slides/scripts/
```

Or clone directly:

```bash
git clone https://github.com/KooBH/frontend-slides-powerpoint.git ~/.claude/skills/powerpoint-slides
```

Then use it by typing `/powerpoint-slides` in Claude Code. Standalone skills are not namespaced.

### Other Coding Agents

Agents such as Codex, Kimi Code, OpenCode, Gemini CLI, or other local coding assistants can use the same core skill. The simplest path is to send the agent this GitHub repo link and ask it to use the PowerPoint Slides skill:

```text
https://github.com/KooBH/frontend-slides-powerpoint
```

If the agent can read GitHub repos or browse files, it should start from `SKILL.md` and load only the referenced support files it needs:

- `STYLE_PRESETS.md`
- `pptx-template.md`
- `animation-patterns.md`
- `bold-template-pack/`
- `scripts/`

Some agents can also install the skill for you if they have filesystem access and a known local skills directory. If not, they can still follow `SKILL.md` directly for the current session.

## Usage

### Create a New Deck

```text
/powerpoint-slides:powerpoint-slides

> "I want to create a pitch deck for my AI startup"
```

If installed manually as a standalone Claude Code skill, use `/powerpoint-slides` instead.

In non-Claude agents, ask the agent to use the PowerPoint Slides skill and point it at this repo or `SKILL.md`.

The skill will:

1. Ask about your content (purpose, length, what you have ready, speaking vs reading deck)
2. Build 3 one-slide style previews, render them to images, and show them to you, inferring the vibe from your brief unless you already named one
3. Let you pick the visual direction
4. Generate the full deck in your chosen style, apply the theme, add motion, validate it, and look at every rendered slide
5. Open it in PowerPoint and tell you how to tweak it (or ask for changes in chat)

### Restyle an Existing PowerPoint

```text
/powerpoint-slides:powerpoint-slides

> "Redesign my quarterly-review.pptx — it looks like every other deck"
```

The skill will:

1. Extract all text, images, tables, chart data, and notes from the original
2. Render the original and show you what it found, for confirmation
3. Let you pick a visual style
4. Generate a new .pptx with all your original content, leaving the original file untouched

### Export

```text
> "Give me a PDF too"
```

`scripts/render-pptx.py deck.pptx --pdf-only` writes a PDF; `--out folder/` writes one PNG per slide.

## Included Styles

### Dark Themes

- **Bold Signal** — Confident, high-impact, vibrant card on dark
- **Electric Studio** — Clean, professional, split-panel
- **Creative Voltage** — Energetic, retro-modern, electric blue + neon
- **Dark Botanical** — Elegant, sophisticated, warm accents

### Light Themes

- **Notebook Tabs** — Editorial, organized, paper with colorful tabs
- **Pastel Geometry** — Friendly, approachable, vertical pills
- **Split Pastel** — Playful, modern, two-color vertical split
- **Vintage Editorial** — Witty, personality-driven, geometric shapes

### Specialty

- **Neon Cyber** — Futuristic, grid-and-glow backgrounds, neon accents
- **Terminal Green** — Developer-focused, scanlines, monospace
- **Swiss Modern** — Minimal, Bauhaus-inspired, geometric
- **Paper & Ink** — Literary, drop caps, pull quotes

### Bold Template Pack

The skill also includes 34 optional bold design systems from
`beautiful-html-templates`, such as **Neo-Grid Bold**, **Editorial Tri-Tone**,
**Creative Mode**, **Broadside**, **Signal**, and **Vellum**.

During style discovery, the preview set is:

- 1 safe preset from `STYLE_PRESETS.md`
- at least 1 bold template option from `bold-template-pack/selection-index.json`
- 1 wildcard option, either another bold template or a self-generated custom design

The agent reads the compact bold template index first, then loads only the
shortlisted candidates' small `preview.md` cards for title-slide previews. It
loads the full `design.md` for exactly one bold template only after the user
picks that template for the final deck. If the user picks a custom wildcard,
the agent expands that preview's own build script and layout system into the full deck.

## Bold Template Gallery

PowerPoint Slides can draw from the 34 bold design systems in [`beautiful-html-templates`](https://github.com/zarazhangrui/beautiful-html-templates). The screenshots below are renders of the source design systems (three per template) and show how each visual system handles different slide layouts; the skill rebuilds the chosen system natively in PowerPoint. Click any template name to inspect the source library.

### [Soft Editorial](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/soft-editorial/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/soft-editorial-4.png" width="32.5%" alt="Soft Editorial — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/soft-editorial-6.png" width="32.5%" alt="Soft Editorial — slide 6" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/soft-editorial-10.png" width="32.5%" alt="Soft Editorial — slide 10" />
</p>

> Cormorant Garamond serif on warm paper with sage, blush, and lemon accents.

### [Editorial Forest](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/editorial-forest/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/editorial-forest-1.png" width="32.5%" alt="Editorial Forest — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/editorial-forest-2.png" width="32.5%" alt="Editorial Forest — slide 2" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/editorial-forest-5.png" width="32.5%" alt="Editorial Forest — slide 5" />
</p>

> Forest green, dusty pink, and warm cream in Source Serif 4 — quiet, intentional quarterly-review aesthetic.

### [Pin & Paper](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/pin-and-paper/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/pin-and-paper-1.png" width="32.5%" alt="Pin & Paper — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/pin-and-paper-11.png" width="32.5%" alt="Pin & Paper — slide 11" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/pin-and-paper-3.png" width="32.5%" alt="Pin & Paper — slide 3" />
</p>

> Yellow paper with safety-pin illustrations, ink-blue handwritten Caveat, paper-grain texture.

### [Sakura Chroma](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/sakura-chroma/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/sakura-chroma-1.png" width="32.5%" alt="Sakura Chroma — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/sakura-chroma-3.png" width="32.5%" alt="Sakura Chroma — slide 3" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/sakura-chroma-4.png" width="32.5%" alt="Sakura Chroma — slide 4" />
</p>

> Vintage Japanese cassette-package aesthetic: cream paper, diagonal rainbow ribbons, condensed bold type, JIS-style spec checkboxes.

### [Stencil & Tablet](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/stencil-tablet/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/stencil-tablet-1.png" width="32.5%" alt="Stencil & Tablet — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/stencil-tablet-3.png" width="32.5%" alt="Stencil & Tablet — slide 3" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/stencil-tablet-8.png" width="32.5%" alt="Stencil & Tablet — slide 8" />
</p>

> Bone paper with stencil-cut headlines and a six-color earth palette: archaeology meets brand.

### [Cobalt Grid](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/cobalt-grid/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/cobalt-grid-1.png" width="32.5%" alt="Cobalt Grid — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/cobalt-grid-3.png" width="32.5%" alt="Cobalt Grid — slide 3" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/cobalt-grid-5.png" width="32.5%" alt="Cobalt Grid — slide 5" />
</p>

> Electric cobalt italic serifs on a graph-paper canvas, anchored by stair-stepped pixel-glitch decorations and slim hairline rules.

### [Vellum](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/vellum/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/vellum-1.png" width="32.5%" alt="Vellum — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/vellum-4.png" width="32.5%" alt="Vellum — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/vellum-8.png" width="32.5%" alt="Vellum — slide 8" />
</p>

> Deep navy canvas with warm-yellow italic Cormorant serifs and a single dusty teal accent. A quiet, scholarly aesthetic.

### [Emerald Editorial](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/emerald-editorial/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/emerald-editorial-1.png" width="32.5%" alt="Emerald Editorial — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/emerald-editorial-3.png" width="32.5%" alt="Emerald Editorial — slide 3" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/emerald-editorial-6.png" width="32.5%" alt="Emerald Editorial — slide 6" />
</p>

> Magazine-cover business deck: emerald + navy + paper with double-rule masthead ornaments and a heavy Bodoni-style display serif.

### [Neo-Grid Bold](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/neo-grid-bold/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/neo-grid-bold-1.png" width="32.5%" alt="Neo-Grid Bold — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/neo-grid-bold-3.png" width="32.5%" alt="Neo-Grid Bold — slide 3" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/neo-grid-bold-8.png" width="32.5%" alt="Neo-Grid Bold — slide 8" />
</p>

> Editorial neo-brutalism with a single neon yellow accent on off-white paper.

### [Editorial Tri-Tone](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/editorial-tri-tone/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/editorial-tri-tone-1.png" width="32.5%" alt="Editorial Tri-Tone — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/editorial-tri-tone-4.png" width="32.5%" alt="Editorial Tri-Tone — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/editorial-tri-tone-3.png" width="32.5%" alt="Editorial Tri-Tone — slide 3" />
</p>

> Three-color editorial system: dusty pink, mustard cream, and deep burgundy, set in Bricolage + Instrument Serif.

### [Creative Mode](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/creative-mode/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/creative-mode-1.png" width="32.5%" alt="Creative Mode — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/creative-mode-4.png" width="32.5%" alt="Creative Mode — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/creative-mode-6.png" width="32.5%" alt="Creative Mode — slide 6" />
</p>

> Cream paper canvas with confident multi-color (green, pink, orange, yellow) accents and Archivo Black display.

### [Monochrome](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/monochrome/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/monochrome-1.png" width="32.5%" alt="Monochrome — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/monochrome-4.png" width="32.5%" alt="Monochrome — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/monochrome-12.png" width="32.5%" alt="Monochrome — slide 12" />
</p>

> Ivory ledger paper with all-black type; Lora serif headlines, Jost body, no color at all.

### [People's Platform (Block & Bold)](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/peoples-platform/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/peoples-platform-1.png" width="32.5%" alt="People's Platform (Block & Bold) — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/peoples-platform-4.png" width="32.5%" alt="People's Platform (Block & Bold) — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/peoples-platform-8.png" width="32.5%" alt="People's Platform (Block & Bold) — slide 8" />
</p>

> Activist poster energy: blue, orange, red on cream, with Alfa Slab + Caveat Brush.

### [Pink Script — After Hours](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/pink-script/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/pink-script-1.png" width="32.5%" alt="Pink Script — After Hours — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/pink-script-4.png" width="32.5%" alt="Pink Script — After Hours — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/pink-script-8.png" width="32.5%" alt="Pink Script — After Hours — slide 8" />
</p>

> Black canvas, hot pink accent, pearl-cream paper, Instrument Serif headlines: late-night editorial luxury.

### [8-Bit Orbit](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/8-bit-orbit/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/8-bit-orbit-1.png" width="32.5%" alt="8-Bit Orbit — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/8-bit-orbit-6.png" width="32.5%" alt="8-Bit Orbit — slide 6" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/8-bit-orbit-5.png" width="32.5%" alt="8-Bit Orbit — slide 5" />
</p>

> Pixel-art neon arcade aesthetic on a deep navy void.

### [BlockFrame](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/block-frame/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/block-frame-1.png" width="32.5%" alt="BlockFrame — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/block-frame-4.png" width="32.5%" alt="BlockFrame — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/block-frame-8.png" width="32.5%" alt="BlockFrame — slide 8" />
</p>

> Neobrutalist deck with pastel-neon color blocks and chunky black borders.

### [Blue Professional](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/blue-professional/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/blue-professional-1.png" width="32.5%" alt="Blue Professional — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/blue-professional-6.png" width="32.5%" alt="Blue Professional — slide 6" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/blue-professional-8.png" width="32.5%" alt="Blue Professional — slide 8" />
</p>

> Cream paper background with electric cobalt blue accents; clean modern professional.

### [Bold Poster](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/bold-poster/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/bold-poster-1.png" width="32.5%" alt="Bold Poster — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/bold-poster-4.png" width="32.5%" alt="Bold Poster — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/bold-poster-8.png" width="32.5%" alt="Bold Poster — slide 8" />
</p>

> Editorial poster aesthetic with massive Shrikhand display and a single fire-engine red accent.

### [Broadside](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/broadside/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/broadside-1.png" width="32.5%" alt="Broadside — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/broadside-4.png" width="32.5%" alt="Broadside — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/broadside-13.png" width="32.5%" alt="Broadside — slide 13" />
</p>

> Dark editorial canvas with a single fire orange accent and bilingual Latin/Chinese type stack.

### [Capsule](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/capsule/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/capsule-1.png" width="32.5%" alt="Capsule — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/capsule-4.png" width="32.5%" alt="Capsule — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/capsule-8.png" width="32.5%" alt="Capsule — slide 8" />
</p>

> Modular pill-shaped cards on warm bone with a full pastel-pop palette.

### [Cartesian](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/cartesian/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/cartesian-1.png" width="32.5%" alt="Cartesian — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/cartesian-4.png" width="32.5%" alt="Cartesian — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/cartesian-8.png" width="32.5%" alt="Cartesian — slide 8" />
</p>

> Quiet warm-neutral palette with classical Playfair serifs; tasteful and unhurried.

### [Coral](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/coral/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/coral-1.png" width="32.5%" alt="Coral — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/coral-4.png" width="32.5%" alt="Coral — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/coral-8.png" width="32.5%" alt="Coral — slide 8" />
</p>

> Cream and coral on near-black, set in oversized Bebas Neue.

### [Daisy Days](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/daisy-days/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/daisy-days-1.png" width="32.5%" alt="Daisy Days — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/daisy-days-4.png" width="32.5%" alt="Daisy Days — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/daisy-days-8.png" width="32.5%" alt="Daisy Days — slide 8" />
</p>

> Cheerful pastel deck with hand-drawn daisies, stars, and rainbows. Friendly, soft, and warm.

### [Grove](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/grove/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/grove-1.png" width="32.5%" alt="Grove — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/grove-4.png" width="32.5%" alt="Grove — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/grove-8.png" width="32.5%" alt="Grove — slide 8" />
</p>

> Forest-green canvas with cream type, classical Playfair serifs, and a single rust accent.

### [Mat](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/mat/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/mat-1.png" width="32.5%" alt="Mat — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/mat-4.png" width="32.5%" alt="Mat — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/mat-8.png" width="32.5%" alt="Mat — slide 8" />
</p>

> Dark sage canvas with bone paper and burnt-orange accent; mid-century modern with wood undertones.

### [Playful](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/playful/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/playful-1.png" width="32.5%" alt="Playful — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/playful-6.png" width="32.5%" alt="Playful — slide 6" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/playful-8.png" width="32.5%" alt="Playful — slide 8" />
</p>

> Sun-warm peach background with Syne display: a friendly indie launch deck.

### [Raw Grid](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/raw-grid/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/raw-grid-1.png" width="32.5%" alt="Raw Grid — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/raw-grid-4.png" width="32.5%" alt="Raw Grid — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/raw-grid-8.png" width="32.5%" alt="Raw Grid — slide 8" />
</p>

> Neo-brutalist deck with thick borders, offset shadows, and a pink/sage/ink palette.

### [Retro Windows](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/retro-windows/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/retro-windows-1.png" width="32.5%" alt="Retro Windows — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/retro-windows-4.png" width="32.5%" alt="Retro Windows — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/retro-windows-8.png" width="32.5%" alt="Retro Windows — slide 8" />
</p>

> Windows 95 chrome: gray title bars, MS Sans Serif, pixel typography, full nostalgia.

### [Retro Zine](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/retro-zine/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/retro-zine-1.png" width="32.5%" alt="Retro Zine — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/retro-zine-4.png" width="32.5%" alt="Retro Zine — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/retro-zine-8.png" width="32.5%" alt="Retro Zine — slide 8" />
</p>

> Beige paper with green accent and Bebas Neue + Caveat: a riso-printed zine in HTML form.

### [Scatterbrain](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/scatterbrain/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/scatterbrain-1.png" width="32.5%" alt="Scatterbrain — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/scatterbrain-4.png" width="32.5%" alt="Scatterbrain — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/scatterbrain-8.png" width="32.5%" alt="Scatterbrain — slide 8" />
</p>

> Post-it inspired: pastel sticky notes, Caveat handwriting, Shrikhand and Zilla Slab type stack.

### [Signal](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/signal/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/signal-1.png" width="32.5%" alt="Signal — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/signal-18.png" width="32.5%" alt="Signal — slide 18" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/signal-8.png" width="32.5%" alt="Signal — slide 8" />
</p>

> Deep navy canvas with bone paper and a single muted-gold accent; institutional with quiet weight.

### [Studio](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/studio/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/studio-1.png" width="32.5%" alt="Studio — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/studio-4.png" width="32.5%" alt="Studio — slide 4" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/studio-8.png" width="32.5%" alt="Studio — slide 8" />
</p>

> Black canvas with electric-yellow type; high-voltage design studio aesthetic.

### [Biennale Yellow](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/biennale-yellow/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/biennale-yellow-1.png" width="32.5%" alt="Biennale Yellow — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/biennale-yellow-5.png" width="32.5%" alt="Biennale Yellow — slide 5" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/biennale-yellow-8.png" width="32.5%" alt="Biennale Yellow — slide 8" />
</p>

> Solar yellow on warm parchment with deep indigo serif and atmospheric sun-glow gradients. Dutch-editorial poster energy.

### [Long Table](https://github.com/zarazhangrui/beautiful-html-templates/tree/main/templates/long-table/)

<p>
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/long-table-1.png" width="32.5%" alt="Long Table — slide 1" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/long-table-3.png" width="32.5%" alt="Long Table — slide 3" />
  <img src="https://raw.githubusercontent.com/zarazhangrui/beautiful-html-templates/main/screenshots/long-table-7.png" width="32.5%" alt="Long Table — slide 7" />
</p>

> Warm cream and rust-red supper-club aesthetic with bold uppercase grotesk headlines, italic Fraunces, and pill-shaped outlined buttons.

## How It Works

Every deck is built by one Node script and finished by a few Python helpers:

```
build-deck.js  --pptxgenjs-->  deck.pptx
                                  |  scripts/apply-theme.py      theme colors + fonts into the file
                                  |  scripts/add-animations.py   transitions + staggered entrance (native, editable)
                                  |  scripts/check-pptx.py       structural validation + outline
                                  v  scripts/render-pptx.py      PNG per slide / PDF, for QA and previews
```

| Script | Purpose |
| --- | --- |
| `scripts/apply-theme.py` | Write a `THEME` (12 scheme colors, heading/body/East-Asian fonts) into `ppt/theme/theme1.xml` so scheme colors resolve correctly and PowerPoint's Design tab can recolor the deck |
| `scripts/add-animations.py` | Add `fade` / `push` / `wipe` / `cover` / `split` / `morph` transitions and `rise` / `fade` / `wipe` / `fly` / `zoom` entrance animations, auto-staggered or click-by-click, per slide or for the whole deck |
| `scripts/make-background.py` | Render gradients, soft glows, grids, dot patterns, film grain, scanlines, and vignettes to an image for layout backgrounds (pptxgenjs has no gradient fills) |
| `scripts/check-pptx.py` | Well-formed XML, relationship targets, content types, chart axis declarations, shapes past the slide edge, probable text overflow, leftover placeholder text, empty slides; plus a per-slide outline with notes |
| `scripts/render-pptx.py` | PowerPoint COM (Windows) or LibreOffice → PNG per slide, PDF, and a labeled contact sheet |
| `scripts/list-fonts.py` | Which font families are installed; check specific names |
| `scripts/extract-pptx.py` | Text (with levels), images (deduplicated), tables, chart series, notes, layout names, and bounding boxes from an existing .pptx |

All scripts use the Python standard library except `extract-pptx.py` (python-pptx) and `make-background.py` and the contact sheet (Pillow).

## Architecture

This skill uses **progressive disclosure** — the main `SKILL.md` is a workflow map, with supporting files loaded on demand only when needed:

| File                      | Purpose                        | Loaded When               |
| ------------------------- | ------------------------------ | ------------------------- |
| `SKILL.md`                | Core workflow and rules        | Always (skill invocation) |
| `STYLE_PRESETS.md`        | 12 curated visual presets with theme colors, fonts, backgrounds | Phase 2 (style selection) |
| `bold-template-pack/selection-index.json` | Compact bold template metadata for candidate selection | Phase 2 (style selection) |
| `bold-template-pack/templates/*/preview.md` | Tiny style cards for shortlisted bold previews | Phase 2 after shortlisting |
| `bold-template-pack/templates/*/design.md` | Full design system for the selected bold template, with a PowerPoint translation policy | Phase 3 after user selection |
| `pptx-template.md`        | Build-script architecture, layouts, components, pptxgenjs gotchas | Phase 3 (generation) |
| `animation-patterns.md`   | Transitions and entrance animations: effect-to-feeling guide | Phase 3 (generation) |
| `scripts/*.py`            | Theme, animation, background, validation, rendering, font, and extraction helpers | Phases 2–6 |

## Philosophy

1. **You don't need to be a designer to make beautiful things.** You just need to react to what you see.
2. **The deck belongs to the user, not to the tool.** A native .pptx with real layouts and theme colors can be edited by anyone, forever, without this skill.
3. **Generic is forgettable.** Every deck should feel custom-crafted for its topic, not template-generated.
4. **Look before you ship.** The agent renders and inspects every slide; overflow and overlap never reach the user.

## Credits

Original skill, workflow, and design philosophy by [@zarazhangrui](https://github.com/zarazhangrui) ([frontend-slides](https://github.com/zarazhangrui/frontend-slides)). Bold template pack derived from [beautiful-html-templates](https://github.com/zarazhangrui/beautiful-html-templates). PowerPoint conversion by [@KooBH](https://github.com/KooBH).

## License

MIT — Use it, modify it, share it.
