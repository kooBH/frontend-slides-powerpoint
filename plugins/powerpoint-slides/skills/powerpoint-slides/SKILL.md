---
name: powerpoint-slides
description: Create distinctive, fully editable PowerPoint (.pptx) decks from scratch or by restyling an existing .pptx. Use when the user wants to build a presentation, make slides for a talk/pitch/report, or redesign an old PowerPoint. Helps non-designers discover their aesthetic through rendered visual previews rather than abstract choices.
---

# PowerPoint Slides

Create native, editable PowerPoint decks with real theme colors, slide layouts, placeholders, speaker notes, charts, transitions, and entrance animations. The output opens in PowerPoint, Keynote, Google Slides, and LibreOffice and can be edited by anyone without touching code.

## Core Principles

1. **Native .pptx, generated from a script** — Every deck is produced by a `build-deck.js` script (pptxgenjs) kept next to the output, then finished by the helper scripts in `scripts/`. Regenerate by re-running the script; never hand-edit XML unless a script cannot do the job.
2. **Show, Don't Tell** — Render candidate title slides to images and let the user react. People discover what they want by seeing it.
3. **Distinctive Design** — No generic "AI slop". Every deck must feel custom-crafted for its topic.
4. **Progressive Disclosure** — Read lightweight style indexes first. For bold templates, use small preview cards for the previews and load the full `design.md` only after the user picks that template.
5. **Fixed 16:9 canvas (NON-NEGOTIABLE)** — Every deck uses `LAYOUT_WIDE`: 13.333 × 7.5 inches. All coordinates are in inches on that canvas. Nothing may extend past the slide edge, and nothing may rely on auto-fit shrinking to hide overflow.
6. **Structured, not painted** — Theme colors, named layouts, and placeholders make the deck restylable from PowerPoint's Design tab. Shapes and text are real objects, never screenshots of a design.

## Design Aesthetics

You tend to converge toward generic, "on distribution" outputs. In slide design this is the "default PowerPoint" look: Calibri on white, a blue title bar, three equal bullet columns, a stock gradient. Avoid it. Make creative, distinctive decks that surprise and delight.

Focus on:

- **Typography:** Choose fonts that are beautiful, unique, and interesting. Avoid Calibri, Arial, Aptos, and Inter as display faces. Pair a characterful display font with a quiet body font. Check what is installed with `python scripts/list-fonts.py` and follow the font policy below.
- **Color & Theme:** Commit to a cohesive palette and write it into the theme so the whole deck follows it. Dominant colors with sharp accents outperform timid, evenly distributed palettes. Draw from editorial design, IDE themes, and cultural aesthetics.
- **Motion:** Use `scripts/add-animations.py` for one well-orchestrated entrance per slide (staggered reveal) and a consistent transition. One good entrance beats scattered effects.
- **Backgrounds:** Create atmosphere and depth rather than defaulting to flat white. Use `scripts/make-background.py` for gradients, soft glows, grids, grain, and scanlines, and apply the result as a layout background.

Avoid generic AI-generated aesthetics:

- Overused fonts (Calibri, Arial, Aptos, Inter, Roboto, system fonts) as display faces
- Cliched color schemes (particularly purple gradients on white, or "corporate blue")
- Accent lines under titles, colored header/footer bars, sidebar stripes, and edge-stripes on cards
- Predictable layouts: title + three bullets, identical card grids on every slide
- Cookie-cutter decks that lack context-specific character

Interpret creatively and make unexpected choices that feel genuinely designed for the context. Vary between light and dark themes, different fonts, different aesthetics. You still tend to converge on common choices (Space Grotesk, Montserrat, for example) across generations. Avoid this: it is critical that you think outside the box.

## Canvas Rules

These invariants apply to EVERY slide in EVERY deck:

- `pres.layout = 'LAYOUT_WIDE'` (13.333 × 7.5 in). Set it before adding anything.
- Author all positions and sizes in inches. Keep a 0.5 in minimum margin from every edge (0.6–0.8 in is the usual rhythm).
- No object may extend past the slide edge unless it is a deliberate full-bleed background panel or image.
- Text must fit its box at the authored size. Never rely on `fit: 'shrink'` to hide overflow. If content exceeds the box, cut words, enlarge the box, or split the slide.
- Minimum sizes: body 14 pt, captions and chrome 10 pt. Titles 36 pt or larger on content slides, 54 pt or larger on title slides.
- Every slide carries a `sectionTitle`, and every content region is a named placeholder or an `objectName`-tagged shape so the user can find things in PowerPoint's Selection Pane.
- Speaker notes go in `slide.addNotes(...)`, never in a text box.
- Translate design-system measurements written for a 1920×1080 web stage with: px ÷ 144 = in, px ÷ 2 = pt, 1 vw = 0.133 in = 9.6 pt, 1 vh = 0.075 in = 5.4 pt.

**When generating, read `pptx-template.md` and follow its build-script architecture and gotchas.**

### Content Density Modes

Ask the user whether this is primarily a reading deck or a speaking deck, then design around that answer:

| Density mode | Best for | Design behavior |
| ------------- | -------- | --------------- |
| **Low density / speaker-led** | Public talks, keynote-style sharing, live explanation | One idea per slide, large type (body 22–28 pt), strong visual hierarchy, generous negative space, 1–3 bullets max, more slides if needed, entrance animations on |
| **High density / reading-first** | Reports, handouts, async review, detailed internal docs | More self-contained slides, structured grids/tables/annotations, 4–8 bullets or 4–6 cards when readable (body 14–18 pt), tighter but still intentional spacing, transitions only |

Baseline limits still apply: no overflow, no overlapping objects, and no text below the minimum sizes. If content exceeds the selected density mode, split it into more slides instead of shrinking until it becomes cramped.

### Font Policy

PowerPoint renders font *names*; a font that is not installed on the viewing machine is silently replaced by Calibri or Arial.

1. Before Phase 2, run `python scripts/list-fonts.py` once and keep the list in mind.
2. Every preset and bold template names a distinctive pair **and** an Office-safe fallback pair. Prefer the distinctive pair when it is installed. If it is not, tell the user in one line that they can install it (Google Fonts / Fontshare, free) for the full effect, and build the previews with whatever is actually installed so the rendered images are honest.
3. If the deck will be presented from another machine or shared widely, say so in delivery: install the fonts there, or in PowerPoint use File → Options → Save → "Embed fonts in the file".
4. Never default to Aptos. Never use more than two families plus an optional mono/label face.
5. For CJK content, set `eaFontFace` in the theme (for example Noto Sans KR / Noto Sans SC / Noto Sans JP) and keep letter spacing at 0 on CJK runs.

---

## Phase 0: Detect Mode

Determine what the user wants:

- **Mode A: New Deck** — Create from scratch. Go to Phase 1.
- **Mode B: Restyle an existing PowerPoint** — Rebuild a .pptx in a new design. Go to Phase 4.
- **Mode C: Enhancement** — Improve a deck this skill already generated (or any .pptx). **Follow Mode C modification rules below.**

### Mode C: Modification Rules

If `build-deck.js` exists next to the deck, edit the script and regenerate; that keeps the deck consistent. Only edit the .pptx directly (python-pptx or XML) when there is no script or the user edited the file in PowerPoint since generation and wants those edits kept.

When enhancing existing decks, canvas fit is the biggest risk:

1. **Before adding content:** Count existing elements on the slide and check against density limits.
2. **Adding images:** Fit them inside the canvas with the standard margins. If the slide already has maximum content, split into two slides.
3. **Adding text:** Max 4–6 bullets per slide. Exceeds limits? Split into continuation slides.
4. **After ANY modification, verify:** run `check-pptx.py`, render with `render-pptx.py`, and look at every changed slide for overflow, overlap, and misalignment.
5. **Proactively reorganize:** If modifications will cause overflow, split content and tell the user. Don't wait to be asked.

---

## Phase 1: Content Discovery (New Decks)

**Ask ALL questions together** so the user fills everything out at once. If the current environment provides a native structured-question UI, use it; otherwise ask in one concise message with clearly numbered choices:

**Question 1 — Purpose** (header: "Purpose"):
What is this presentation for? Options: Pitch deck / Teaching-Tutorial / Conference talk / Internal presentation

**Question 2 — Length** (header: "Length"):
Approximately how many slides? Options: Short 5–10 / Medium 10–20 / Long 20+

**Question 3 — Content** (header: "Content"):
Do you have content ready? Options: All content ready / Rough notes / Topic only

**Question 4 — Density** (header: "Density"):
How dense should the deck feel? Options:

- "Low density / speaker-led" — Big ideas, fewer words, more visual breathing room
- "High density / reading-first" — More self-contained detail for async reading

Remember the user's density choice. It affects slide count, typography scale, amount of text per slide, layout density, whether entrance animations are on by default, and whether to favor cinematic presenter slides or self-contained reading slides.

If the user has content, ask them to share it (text, notes, an outline, a folder of images, or a .docx/.md file).

### Step 1.2: Image Evaluation (if images provided)

If the user has no images → skip to Phase 2. Shapes, charts, icons, and generated backgrounds carry the visual weight; this is a fully supported first-class path.

If the user provides an image folder:

1. **Scan** — List all image files (.png, .jpg, .svg, .webp, etc.). Note that pptxgenjs embeds PNG/JPG/GIF directly; convert SVG to PNG first (Pillow with cairosvg, or any converter) unless the SVG is simple.
2. **Inspect each image** — Use the agent's available image-understanding capability. If image reading is unavailable, use filenames/metadata and ask the user to clarify only when needed.
3. **Evaluate** — For each: what it shows, USABLE or NOT USABLE (with reason), what concept it represents, dominant colors.
4. **Co-design the outline** — Curated images inform slide structure alongside text. This is NOT "plan slides then add images": design around both from the start (e.g., 3 screenshots → 3 feature slides, 1 logo → title/closing slide).
5. **Confirm the outline** using the same structured-question mechanism when available: "Does this slide outline and image selection look right?" Options: Looks good / Adjust images / Adjust outline

**Logo in previews:** If a usable logo was identified, place it on each style preview in Phase 2 so the user sees their brand styled three different ways.

---

## Phase 2: Style Discovery

**This is the "show, don't tell" phase.** Most people can't articulate design preferences in words.

### Step 2.0: Generate 3 Style Previews Directly

Based on purpose, audience, mood, and content density, build 3 distinct one-slide .pptx files showing typography, colors, background treatment, and overall aesthetic, render them to PNG, and show the images.

Do not ask the user whether they want options or a preset picker. The default discovery experience is always visual comparison.

If the user already gave a vibe, use it. If they did not, infer the likely mood from the occasion, audience, content, and stakes. Keep the options diverse enough that the user can react visually instead of needing to articulate taste up front.

If the user explicitly names a preset or bold template, honor that as one option and generate the remaining preview slots around it.

Read [STYLE_PRESETS.md](STYLE_PRESETS.md) for safe preset candidates. If [bold-template-pack/selection-index.json](bold-template-pack/selection-index.json) exists, read that compact index too, but do not read any `design.md` files yet.

| Mood                | Suggested Presets                                  |
| ------------------- | -------------------------------------------------- |
| Impressed/Confident | Bold Signal, Electric Studio, Dark Botanical       |
| Excited/Energized   | Creative Voltage, Neon Cyber, Split Pastel         |
| Calm/Focused        | Notebook Tabs, Paper & Ink, Swiss Modern           |
| Inspired/Moved      | Dark Botanical, Vintage Editorial, Pastel Geometry |

**Preview mix rules:**

- Generate 3 previews by default: 1 safe preset from `STYLE_PRESETS.md`, at least 1 bold template from `bold-template-pack/selection-index.json`, and 1 wildcard.
- The wildcard may be either a second bold template or a self-generated custom design. Choose whichever creates the strongest, most useful contrast for the user's occasion, audience, mood, and content.
- Do not force every expressive option to come from the template library. If the brief has a sharper, more specific design opportunity than the available templates, use the wildcard slot to design freely.
- For conservative or high-stakes decks, make the safe preset especially restrained; choose a calm, higher-formality bold template; make the wildcard either another restrained template or a custom design that feels authoritative rather than decorative.
- For expressive decks, keep the safe preset as a readable fallback; choose one strong bold template; make the wildcard adventurous, context-specific, and clearly different from both other previews.
- If bold template matches feel weak, use the wildcard as a custom design or fall back to another safe preset instead of forcing a template.

**Custom wildcard design rules:**

- Follow the Design Aesthetics section above: no generic "AI slop", no default font/color/layout choices, no purple-gradient-on-white clichés, no cookie-cutter dashboard/card look.
- Match the user's stated occasion, audience, mood/vibe, and content density. The custom design should feel authored for this deck, not merely "stylish."
- Make a deliberate visual thesis: distinctive typography, a committed palette, a recognizable layout system, and one strong atmospheric or graphic device (a generated background, a signature shape, a chrome convention).
- Keep it feasible for a full deck. The preview must imply a design system that can expand into section, content, quote, comparison, chart, and closing slides.
- Never render "custom", "wildcard", "AI-generated", or design-process labels on the slide itself.

**Bold template selection rules:**

- Match user purpose and mood against `mood`, `tone`, `best_for`, `avoid_for`, `formality`, `density`, and `scheme`.
- Treat `best_for` examples as soft signals, not strict industry filters.
- Keep the three previews genuinely different from each other.
- After choosing bold template candidate(s), read only those candidate(s)' `preview.md` files from the `preview_md` paths in the selection index.
- Use `preview.md` only for title-slide previews. Do not read full `design.md` files until the user picks the final template.

**Preview authenticity rules (NON-NEGOTIABLE):**

- Every style preview must look like a real first slide from the user's deck, not a diagnostic card.
- Never render internal workflow text on a slide: no `preview`, `generated from`, `preview.md`, `template`, `preset`, `style option`, `Option A/B/C`, file names, paths, or source-doc labels.
- Never render template names or slug names on the slide itself. Template/style names belong only in the message to the user.
- Never render user requirement notes as slide content, such as "sharp and provocative", "safe option", "bold option", "for internal sharing", or "audience: ...", unless the user explicitly wants that exact phrase to appear in the deck.
- If the slide needs chrome, use real deck chrome only: the deck title, section title, date, author, company, page number, or a genuine content phrase from the user's material.
- Before showing previews, read the rendered image and revise if any internal metadata appears, if a font fell back to Calibri/Arial unexpectedly, or if text overflows.

**How to build the previews:**

1. Create `.powerpoint-slides/previews/` and write one small pptxgenjs script per style (`style-a.js`, `style-b.js`, `style-c.js`), each producing a single title slide at `LAYOUT_WIDE` with the real deck title, subtitle, author/date. Use the same THEME + layout approach as the final deck so the preview is an honest sample.
2. Run them, then `python scripts/apply-theme.py style-a.pptx theme-a.json` (and so on).
3. Render: `python scripts/render-pptx.py style-a.pptx --out .powerpoint-slides/previews/a` (repeat for b and c). Use `--width 1280` to keep the images light.
4. Look at each PNG yourself first (authenticity check), then show all three images to the user in one message, labelled only with the style names.

### Step 2.1: User Picks

Ask (header: "Style"):
Which style preview do you prefer? Options: Style A: [Name] / Style B: [Name] / Style C: [Name] / Mix elements

If "Mix elements", ask for specifics.

---

## Phase 3: Generate Deck

Generate the full deck using content from Phase 1 (text, or text + curated images) and style from Phase 2.

Apply the user's density choice throughout the deck:

- **Low density / speaker-led:** Use more slides with fewer ideas per slide. Favor large headings, short phrases, visual metaphors, section beats, quote/statement slides, and presenter-friendly pacing. Entrance animations on (auto, staggered).
- **High density / reading-first:** Make slides more self-contained. Use structured grids, comparison tables, native charts, annotated diagrams, captions, and concise explanatory copy. Keep hierarchy strong so it feels designed, not like a document pasted onto slides. Transitions only; no entrance animations unless asked.

If the user's stated needs are mixed, choose the closer of the two modes instead of inventing a middle option: live audience persuasion defaults low-density; async circulation or detailed review defaults high-density.

Never let high density become visual clutter. If a high-density slide starts to overflow, split it or redesign it into a clearer structure.

If the user selected a bold template from `bold-template-pack`, read that one template's full `design.md` before generating. Do not read the other bold templates. Treat `design.md` as the design recipe:

- Preserve its fonts, palette, decorative vocabulary, spacing rhythm, and component grammar.
- Translate its CSS measurements into inches and points with the conversion rule in Canvas Rules; the `PowerPoint Translation Policy` section at the top of each `design.md` explains how each CSS device maps to a PowerPoint construct.
- Do not copy demo slide content or mimic the source template too literally.
- Do not read or copy `template.html` from the source library.

If the user selected a self-generated custom wildcard, treat that preview's script as the design recipe:

- Preserve its fonts, palette, decorative vocabulary, spacing rhythm, grid logic, and component grammar.
- Expand the same visual system across the full deck. Do not switch to a preset or bold template after the user has chosen the custom direction.
- Design any missing slide layouts from that system rather than importing patterns from another style.

**Before generating, read these supporting files:**

- [pptx-template.md](pptx-template.md) — build-script architecture, layouts, components, pptxgenjs gotchas
- [animation-patterns.md](animation-patterns.md) — motion reference for the chosen feeling

**Build sequence (always in this order):**

```bash
npm install pptxgenjs            # once per project folder, if `require('pptxgenjs')` fails
node build-deck.js               # writes <deck>.pptx
python scripts/apply-theme.py <deck>.pptx theme.json
python scripts/add-animations.py <deck>.pptx [--entrance rise|fade|none] [--transition fade|push|morph]
python scripts/check-pptx.py <deck>.pptx
python scripts/render-pptx.py <deck>.pptx --grid
```

Then do visual QA (Phase 5) before delivering.

**Key requirements:**

- One `build-deck.js` + `theme.json` + any generated background PNGs saved next to the deck, with clear `// === SECTION ===` comments so the user (or a later session) can regenerate and tweak.
- Theme colors and fonts written into the file with `apply-theme.py`; slide text uses scheme colors so PowerPoint's Design → Variants can recolor the deck.
- Named layouts for every slide frame used (title, section, content, two-column, stat, quote, chart, closing, ...), with placeholders for title and body regions.
- Every slide has speaker notes (even one line) and a section.
- Charts are native (`addChart`), tables are native (`addTable`), icons are vector shapes or PNGs with transparent backgrounds. Never paste a screenshot where an editable object is possible.

---

## Phase 4: Restyle an Existing PowerPoint

When the user brings a .pptx to redesign:

1. **Extract content** — Run `python scripts/extract-pptx.py <input.pptx> <output_dir>` (install python-pptx if needed: `pip install python-pptx`). It writes `extracted-slides.json` and an `assets/` folder with every image.
2. **Render the original** — `python scripts/render-pptx.py <input.pptx> --grid --width 960` and look at the contact sheet so you understand the existing structure and what the user is used to.
3. **Confirm with user** — Present extracted slide titles, content summaries, image counts, and anything you intend to merge or split.
4. **Style selection** — Proceed to Phase 2 for style discovery.
5. **Generate** — Rebuild through Phase 3 in the chosen style, preserving all text, images (from `assets/`), tables, chart data, slide order, and speaker notes. Keep the original file untouched; write a new file.

---

## Phase 5: Quality Assurance and Delivery

**QA is required, not optional.** Your first render usually has a few real issues.

1. **Structure** — `python scripts/check-pptx.py <deck>.pptx` must report 0 errors. Read its outline to confirm nothing is missing, duplicated, or out of order. Resolve warnings about overflow, empty placeholders, and leftover prompt text.
2. **Render** — `python scripts/render-pptx.py <deck>.pptx --grid`. Look at every slide image fresh (a subagent works well if available). Check first for text overflow or cut-off text, then: overlapping objects, elements closer than 0.3 in, uneven gaps, margins under 0.5 in, misaligned columns, low-contrast text, titles at different heights across slides of the same type, inconsistent chart label styling, words or numbers broken across lines, fonts that fell back unexpectedly.
3. **Fix in the script, rebuild, re-render** only the slides you changed (or all, it is cheap), until clean.
4. **Clean up** — Delete `.powerpoint-slides/previews/` and any `*-render/` folders unless the user wants the images.
5. **Open** — Open the deck (`start deck.pptx` on Windows, `open deck.pptx` on macOS, `xdg-open` on Linux).
6. **Summarize** — Tell the user:
   - File location, style name, slide count, which fonts are used and whether they are installed here
   - That the deck is fully editable: theme colors under Design → Variants, layouts under View → Slide Master, animations in the Animations pane
   - That `build-deck.js` next to the deck regenerates it; ask for revisions in chat or edit directly in PowerPoint
   - Offer the natural post-draft actions: revisions, a PDF export, or handout/notes pages

---

## Phase 6: Share & Export (Optional)

After delivery, **ask the user:** _"Would you like a PDF version for email or printing, or slide images for a doc or chat?"_

Options:

- **Export to PDF** — `python scripts/render-pptx.py <deck>.pptx --pdf-only`. Animations are replaced by their final state; transitions are dropped. Mention this.
- **Export slide images** — `python scripts/render-pptx.py <deck>.pptx --out <folder>` gives one PNG per slide.
- **No thanks**

**⚠ Export gotchas:**

- On Windows with Office, rendering uses PowerPoint itself (pixel-accurate). On macOS/Linux it uses LibreOffice; fonts not installed there are substituted and the PNGs may differ slightly from PowerPoint.
- Large decks with many full-bleed images produce large PDFs. If the PDF exceeds 10 MB, offer to compress by regenerating backgrounds at 1280×720.
- If the user will present from another machine, repeat the font advice from the Font Policy (install the fonts there or embed them).

---

## Supporting Files

| File                                               | Purpose                                                              | When to Read              |
| -------------------------------------------------- | -------------------------------------------------------------------- | ------------------------- |
| [STYLE_PRESETS.md](STYLE_PRESETS.md)               | 12 curated visual presets with theme colors, fonts, and signature elements | Phase 2 (style selection) |
| [bold-template-pack/selection-index.json](bold-template-pack/selection-index.json) | Compact bold template metadata for candidate selection | Phase 2 (style selection) |
| [bold-template-pack/templates/*/preview.md](bold-template-pack/templates/) | Lightweight style cards for shortlisted bold title previews | Phase 2 after shortlisting |
| [bold-template-pack/templates/*/design.md](bold-template-pack/templates/) | Detailed design-system docs for the selected bold template only | Phase 3 after user selection |
| [pptx-template.md](pptx-template.md)               | Build-script architecture, layouts, components, pptxgenjs gotchas    | Phase 2 and 3 (generation) |
| [animation-patterns.md](animation-patterns.md)     | Transitions and entrance animations: effect-to-feeling guide         | Phase 3 (generation)      |
| [scripts/apply-theme.py](scripts/apply-theme.py)   | Write theme colors and fonts into the .pptx                          | Phase 2, 3                |
| [scripts/add-animations.py](scripts/add-animations.py) | Add transitions and staggered entrance animations                | Phase 3                   |
| [scripts/make-background.py](scripts/make-background.py) | Generate gradient / glow / grid / grain background images       | Phase 2, 3                |
| [scripts/check-pptx.py](scripts/check-pptx.py)     | Structural validation and content outline                            | Phase 3, 5                |
| [scripts/render-pptx.py](scripts/render-pptx.py)   | Render slides to PNG / PDF (PowerPoint COM or LibreOffice)           | Phase 2, 5, 6             |
| [scripts/list-fonts.py](scripts/list-fonts.py)     | Check which fonts are installed                                      | Phase 2                   |
| [scripts/extract-pptx.py](scripts/extract-pptx.py) | Extract text, images, tables, charts, notes from an existing .pptx  | Phase 4 (restyle)         |
