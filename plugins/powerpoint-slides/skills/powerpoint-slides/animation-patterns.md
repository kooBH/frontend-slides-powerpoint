# Animation Patterns Reference

Motion in PowerPoint comes in two layers: **transitions** (between slides) and **entrance animations** (objects appearing on a slide). pptxgenjs writes neither, so the skill adds them after the build with `scripts/add-animations.py`, which injects native, editable PowerPoint animation XML. Everything it adds shows up in PowerPoint's Transitions tab and Animations pane and can be changed by the user.

```bash
python scripts/add-animations.py deck.pptx [--transition ...] [--entrance ...] [--mode auto|click] [--stagger MS]
```

Defaults: `--transition fade --entrance rise --mode auto --duration 500 --stagger 120`. That is one orchestrated, staggered reveal per slide — the PowerPoint equivalent of a well-timed page-load animation.

## Effect-to-Feeling Guide

| Feeling | Transition | Entrance | Timing | Visual cues |
|---------|-----------|----------|--------|-------------|
| **Dramatic / Cinematic** | `fade --speed slow` or `morph` | `fade` | `--duration 900 --stagger 220` | Dark backgrounds, generated glows, full-bleed images with scrim |
| **Techy / Futuristic** | `wipe --direction r` | `wipe` | `--duration 400 --stagger 90` | Grid/scanline backgrounds, mono chrome, neon shadow on key shapes |
| **Playful / Friendly** | `push --direction l` | `zoom` | `--duration 450 --stagger 140` | Rounded shapes, pastel badges, bouncy rhythm |
| **Professional / Corporate** | `fade --speed fast` | `rise` | `--duration 350 --stagger 100` | Navy/slate/charcoal, precise spacing, native charts |
| **Calm / Minimal** | `fade --speed slow` | `fade` | `--duration 800 --stagger 200` | High whitespace, muted palette, serif typography |
| **Editorial / Magazine** | `fade` | `rise` | `--duration 500 --stagger 150` | Strong type hierarchy, pull quotes, grid-breaking layouts |
| **Reading-first deck** | `fade --speed fast` | `none` | — | Nothing moves on the slide; readers page at their own pace |

## Entrance Effects Available

| `--entrance` | What it does | PowerPoint name |
|---|---|---|
| `rise` | fade in while drifting up ~0.3 in — the default "reveal" | Fade (+ custom motion) |
| `fade` | opacity only | Fade |
| `wipe` | reveal from the left edge | Wipe |
| `fly` | slide in from below | Fly In |
| `zoom` | scale up from the center with a fade | Zoom |
| `none` | no entrance animations | — |

`--mode auto` plays all objects automatically when the slide appears, each delayed by `--stagger` from the previous one, in z-order (the order shapes were added in `build-deck.js`: title first, then kicker, then body, then visuals works well).

`--mode click` makes every object its own click step, in z-order. Use it for speaker-led decks where bullets should land one at a time.

Shapes covering 85% or more of the slide (background panels) are skipped so the stage never fades in with the content; pass `--include-large` to override. Slide number, footer, and date placeholders are never animated.

## Transitions Available

`fade`, `push`, `wipe`, `cover`, `split`, `morph`, `none` with `--speed slow|med|fast` and `--direction l|r|u|d` for the directional ones. `--advance-after 8000` turns on auto-advance (kiosk / looping booth decks).

**Morph** moves objects that exist on both slides to their new positions and is the single most cinematic transition available. PowerPoint matches objects by name: give the shapes that should travel the same `objectName` on consecutive slides, prefixed with `!!` (for example `objectName: '!!hero-card'`), and use a `SECTION` → `CONTENT` pair where the same shape shrinks from a hero to a chrome element. Keep the deck openable in older PowerPoint: the script writes a fade fallback automatically.

## Patterns

### One orchestrated reveal per slide (default)

Order the `addText` / `addShape` calls in `build-deck.js` in the sequence you want them to appear: kicker → title → body → visual → caption. Run with defaults. Every slide gets the same rhythm; the deck feels designed, not decorated.

### Bullet-by-bullet on click

A single text box with five paragraphs is one object and animates as one. For click-by-click bullets, add each bullet as its own text box (same x, stepped y), then run `--mode click`. Apply it only to the slides that need it:

```bash
python scripts/add-animations.py deck.pptx --slides 4,7 --mode click --entrance fade
python scripts/add-animations.py deck.pptx --skip 4,7                 # defaults for the rest
```

(The script replaces what it previously added on the slides it touches, so run the specific slides last or use `--skip`.)

### Title slide: nothing moves

Title slides often read better static. `--skip 1`.

### Section beats

Section dividers with `morph` from the previous content slide and a slow `fade` entrance on the big title feel like chapter openings:

```bash
python scripts/add-animations.py deck.pptx --slides 5,12,19 --transition morph --entrance fade --duration 900
```

### Reduced motion

Some audiences and venues need no motion at all. `--clear` removes everything the script added; `--entrance none --transition fade --speed fast` keeps a bare minimum.

## Background Effects

Backgrounds carry "atmosphere" in PowerPoint the way CSS gradients and pseudo-elements do on the web. Generate them once with `scripts/make-background.py` and assign them to layouts (`background: { path }`), so every slide on that layout shares the texture.

```bash
# gradient mesh depth
python scripts/make-background.py bg.png --gradient 0a0f1c,111827 --angle 120 --glow 00ffcc@80%,20%:0.2 --glow 7800ff@20%,80%:0.25
# paper grain
python scripts/make-background.py bg.png --solid faf9f7 --noise 0.03
# structural grid
python scripts/make-background.py bg.png --solid 1c2644 --grid 80:ffffff:0.03
# CRT scanlines
python scripts/make-background.py bg.png --solid 0d1117 --scanlines 4:000000:0.18 --vignette 0.3
```

Keep backgrounds at 1920×1080; the file is embedded once per layout, not once per slide.

## Hover, Tilt, Parallax, Particles

These web effects have no PowerPoint equivalent and are intentionally dropped. Their *job* — depth and delight — is done here by: the generated background, one accent shadow or glow on the hero shape, a consistent entrance rhythm, and Morph between related slides.

## Troubleshooting

| Problem | Fix |
|---------|-----|
| PowerPoint says the file needs repair after animating | Run `check-pptx.py`; if clean, re-run `add-animations.py --clear` then add again. Report the slide number — the deck itself is intact |
| Background panel fades in with the content | It was under 85% of the slide. Make it full-bleed, or re-run with the panel named in `--skip`-free mode after enlarging it |
| Bullets appear all at once | They are one text box. Split into one text box per bullet and use `--mode click` |
| Morph does not move objects | Names must match across the two slides and start with `!!`; both slides need the object |
| Transition looks different in Keynote / Google Slides | Only fade, push, and wipe are universal. Use those for decks that leave PowerPoint |
| PDF export shows everything at once | Expected: PDF captures the final state of each slide |
