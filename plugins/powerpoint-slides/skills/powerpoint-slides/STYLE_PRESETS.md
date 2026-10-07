# Style Presets Reference

Curated visual styles for PowerPoint Slides. Each preset is inspired by real design references — no generic "AI slop" aesthetics. **Abstract shapes only — no illustrations.**

Each preset gives:

- **Theme** — the `THEME.colors` block for `build-deck.js` / `apply-theme.py`. Convention: `lt1` is the default slide surface and `dk1` the ink on it (even for dark decks); `lt2`/`dk2` are the alternate surface and its ink; `accent1` is *the* accent.
- **Typography** — a distinctive pair (install from Google Fonts / Fontshare) **and** an Office-safe fallback pair that ships with Microsoft Office on Windows and macOS. Check availability with `python scripts/list-fonts.py "<font>"`.
- **Background** — a `scripts/make-background.py` recipe when the look needs atmosphere beyond a flat fill.
- **Signature elements** — translated into PowerPoint constructs (shapes, layouts, chrome).

---

## Dark Themes

### 1. Bold Signal

**Vibe:** Confident, bold, modern, high-impact

**Layout:** Large colored card as the focal panel on a dark gradient. Section number top-left, navigation breadcrumbs top-right, title bottom-left of the card.

**Typography:**
- Display: `Archivo Black` · Body: `Space Grotesk`
- Office-safe fallback: `Arial Black` · `Corbel`

**Theme:**
```json
{ "dk1": "FFFFFF", "lt1": "1A1A1A", "dk2": "1A1A1A", "lt2": "FF5722",
  "accent1": "FF5722", "accent2": "BDBDBD", "accent3": "2D2D2D", "accent4": "3A3A3A", "accent5": "FFB74D", "accent6": "121212" }
```

**Background:** `make-background.py bg.png --gradient 1a1a1a,2d2d2d,1a1a1a --angle 135`

**Signature Elements:**
- One large `roundRect` card (orange/coral) covering ~60% of the slide as the content panel; dark ink on it
- Oversized section numbers (01, 02, …) at 96 pt in the display face, top-left
- Breadcrumb text row top-right in 11 pt with inactive items at `transparency: 60`
- Everything on an 8-column grid (column = 1.47 in)

---

### 2. Electric Studio

**Vibe:** Bold, clean, professional, high contrast

**Layout:** Horizontal split — white top panel, electric blue bottom panel. Brand marks in corners.

**Typography:**
- Display: `Manrope` (800) · Body: `Manrope`
- Office-safe fallback: `Gill Sans MT` · `Gill Sans MT`

**Theme:**
```json
{ "dk1": "0A0A0A", "lt1": "FFFFFF", "dk2": "FFFFFF", "lt2": "4361EE",
  "accent1": "4361EE", "accent2": "0A0A0A", "accent3": "6B7280", "accent4": "E5E7EB", "accent5": "A5B4FC", "accent6": "1E1E2E" }
```

**Signature Elements:**
- Two full-width `rect` panels (top 58% white, bottom 42% blue) defined on the layout
- A single 0.08 in accent bar on the panel edge is the *only* line device (never under titles)
- Quote typography as hero element: 40 pt display text in the blue panel in white
- Minimal, confident spacing: 1 in margins on title slides

---

### 3. Creative Voltage

**Vibe:** Bold, creative, energetic, retro-modern

**Layout:** Vertical split — electric blue left, dark right. Monospace accents.

**Typography:**
- Display: `Syne` (800) · Mono: `Space Mono`
- Office-safe fallback: `Impact` · `Consolas`

**Theme:**
```json
{ "dk1": "FFFFFF", "lt1": "0066FF", "dk2": "FFFFFF", "lt2": "1A1A2E",
  "accent1": "D4FF00", "accent2": "FFFFFF", "accent3": "9DB7FF", "accent4": "2A2A44", "accent5": "FF3DAE", "accent6": "003399" }
```

**Background:** `make-background.py bg-right.png --solid 1a1a2e --dots 28:ffffff:0.10` (halftone feel on the dark panel)

**Signature Elements:**
- Electric blue + neon yellow contrast; yellow is used for 2–3 words per slide, never for paragraphs
- Dot-grid halftone texture on the dark panel (generated PNG)
- Neon badge pills: `roundRect` with `rectRadius: 0.3`, yellow fill, 11 pt mono uppercase label
- One mono-set caption or timestamp per slide as chrome

---

### 4. Dark Botanical

**Vibe:** Elegant, sophisticated, artistic, premium

**Layout:** Centered or left-aligned content on near-black. Soft blurred color blobs in a corner.

**Typography:**
- Display: `Cormorant` (400/600) · Body: `IBM Plex Sans` (300/400)
- Office-safe fallback: `Perpetua` · `Candara`

**Theme:**
```json
{ "dk1": "E8E4DF", "lt1": "0F0F0F", "dk2": "0F0F0F", "lt2": "E8E4DF",
  "accent1": "D4A574", "accent2": "9A9590", "accent3": "E8B4B8", "accent4": "2A2724", "accent5": "C9B896", "accent6": "1A1816" }
```

**Background:** `make-background.py bg.png --solid 0f0f0f --glow e8b4b8@85%,15%:0.30:30% --glow d4a574@75%,30%:0.22:25% --noise 0.03`

**Signature Elements:**
- Abstract soft gradient circles in a corner (generated background, never drawn shapes)
- Warm accents (gold, pink, terracotta) on rules, numerals, italic emphasis
- Thin vertical accent line (0.02 in wide, 1.5 in tall) beside headlines
- Italic display face for one emphasized word per headline, in gold
- **No illustrations — only abstract blurred color**

---

## Light Themes

### 5. Notebook Tabs

**Vibe:** Editorial, organized, elegant, tactile

**Layout:** Cream paper panel on a dark outer surface. Colored section tabs on the right edge.

**Typography:**
- Display: `Bodoni Moda` · Body: `DM Sans`
- Office-safe fallback: `Bodoni MT` · `Tw Cen MT`

**Theme:**
```json
{ "dk1": "1A1A1A", "lt1": "F8F6F1", "dk2": "F8F6F1", "lt2": "2D2D2D",
  "accent1": "98D4BB", "accent2": "C7B8EA", "accent3": "F4B8C5", "accent4": "A8D8EA", "accent5": "FFE6A7", "accent6": "6B6B6B" }
```

**Signature Elements:**
- Layout background `2D2D2D`; a paper `rect` (`F8F6F1`) inset 0.35 in with a soft outer shadow (`blur: 12, offset: 4, opacity: 0.35`)
- Five vertical tabs on the right edge: 0.35 in wide `rect`s in the five accent colors, stacked, with 9 pt rotated labels (`rotate: 270`)
- Binder-hole decorations: three 0.18 in circles down the left margin, fill `2D2D2D`
- Section slides change which tab is "active" (full opacity) — others at `transparency: 40`

---

### 6. Pastel Geometry

**Vibe:** Friendly, organized, modern, approachable

**Layout:** White rounded card on a pastel surface. Vertical pills of varying height on the right edge.

**Typography:**
- Display: `Plus Jakarta Sans` (800) · Body: `Plus Jakarta Sans`
- Office-safe fallback: `Century Gothic` · `Century Gothic`

**Theme:**
```json
{ "dk1": "1F2A33", "lt1": "C8D9E6", "dk2": "1F2A33", "lt2": "FAF9F7",
  "accent1": "7C6AAD", "accent2": "5A7C6A", "accent3": "F0B4D4", "accent4": "A8D4C4", "accent5": "9B8DC4", "accent6": "E4ECF2" }
```

**Signature Elements:**
- `roundRect` card (`rectRadius: 0.25`, fill `FAF9F7`, shadow `blur: 10, offset: 3, opacity: 0.18`) inset 0.5 in
- **Vertical pills on the right edge:** five `roundRect`s, same width (0.3 in), heights short → medium → tall → medium → short, in accent colors
- Small action icon (circle + glyph shape) in the top-right corner of the card
- Generous 0.6 in card padding; body 16–18 pt

---

### 7. Split Pastel

**Vibe:** Playful, modern, friendly, creative

**Layout:** Two-color vertical split (peach left, lavender right).

**Typography:**
- Display: `Outfit` (800) · Body: `Outfit`
- Office-safe fallback: `Tw Cen MT` · `Tw Cen MT`

**Theme:**
```json
{ "dk1": "1A1A1A", "lt1": "F5E6DC", "dk2": "1A1A1A", "lt2": "E4DFF0",
  "accent1": "1A1A1A", "accent2": "C8F0D8", "accent3": "F0F0C8", "accent4": "F0D4E0", "accent5": "8C7BB5", "accent6": "D9C7B8" }
```

**Background:** right panel `make-background.py bg-right.png --solid e4dff0 --grid 48:1a1a1a:0.06`

**Signature Elements:**
- Two full-height `rect` panels on the layout (left 50% peach, right 50% lavender with a faint grid PNG)
- Playful badge pills (`roundRect`, mint/yellow/pink) with a small shape icon and a 12 pt label
- Rounded CTA button: `roundRect` with `rectRadius: 0.4`, black fill, white 14 pt label
- Titles sit on the left panel; visuals, badges, and callouts on the right

---

### 8. Vintage Editorial

**Vibe:** Witty, confident, editorial, personality-driven

**Layout:** Centered content on cream. Abstract geometric shapes as accents.

**Typography:**
- Display: `Fraunces` (700/900) · Body: `Work Sans`
- Office-safe fallback: `Bookman Old Style` · `Candara`

**Theme:**
```json
{ "dk1": "1A1A1A", "lt1": "F5F3EE", "dk2": "F5F3EE", "lt2": "1A1A1A",
  "accent1": "E8D4C0", "accent2": "555555", "accent3": "C9B7A0", "accent4": "DED8CE", "accent5": "B04A2A", "accent6": "2F2A24" }
```

**Signature Elements:**
- Abstract geometric trio on the title and section slides: a 1.2 in circle outline (`line` 0.03 in, no fill), a 2 in horizontal rule, and a 0.25 in filled dot
- Bold bordered CTA boxes: `rect` with 0.04 in black border, no fill, 16 pt uppercase label
- Witty, conversational copy style; short titles that read like magazine deks
- **No illustrations — only geometric shapes**

---

## Specialty Themes

### 9. Neon Cyber

**Vibe:** Futuristic, techy, confident

**Typography:** `Clash Display` + `Satoshi` (Fontshare) · Office-safe fallback: `Franklin Gothic Demi` · `Corbel`

**Theme:**
```json
{ "dk1": "FFFFFF", "lt1": "0A0F1C", "dk2": "0A0F1C", "lt2": "00FFCC",
  "accent1": "00FFCC", "accent2": "FF00AA", "accent3": "7C8DB5", "accent4": "1B2540", "accent5": "4ADEFF", "accent6": "111827" }
```

**Background:** `make-background.py bg.png --solid 0a0f1c --grid 64:00ffcc:0.07 --glow 00ffcc@80%,20%:0.18:30% --glow ff00aa@15%,85%:0.14:25% --vignette 0.35`

**Signature:** Grid + glow background PNG, neon glow via an outer shadow in the accent color (`shadow: { type: 'outer', color: '00FFCC', blur: 18, offset: 0, opacity: 0.6 }`) on key shapes, mono-style 11 pt uppercase chrome, cyan/magenta used only as accents on dark

---

### 10. Terminal Green

**Vibe:** Developer-focused, hacker aesthetic

**Typography:** `JetBrains Mono` (monospace only) · Office-safe fallback: `Consolas`

**Theme:**
```json
{ "dk1": "C9D1D9", "lt1": "0D1117", "dk2": "0D1117", "lt2": "39D353",
  "accent1": "39D353", "accent2": "8B949E", "accent3": "58A6FF", "accent4": "21262D", "accent5": "F78166", "accent6": "161B22" }
```

**Background:** `make-background.py bg.png --solid 0d1117 --scanlines 4:000000:0.18 --vignette 0.25`

**Signature:** Scanline background, a blinking-cursor glyph (`▌`) after titles, code blocks as `rect` panels (`161B22`) with 14 pt mono text and syntax colors as separate runs, a `$` prompt prefix on section titles

---

### 11. Swiss Modern

**Vibe:** Clean, precise, Bauhaus-inspired

**Typography:** `Archivo` (800) + `Nunito` · Office-safe fallback: `Franklin Gothic Demi` · `Trebuchet MS`

**Theme:**
```json
{ "dk1": "000000", "lt1": "FFFFFF", "dk2": "FFFFFF", "lt2": "000000",
  "accent1": "FF3300", "accent2": "4D4D4D", "accent3": "9A9A9A", "accent4": "E6E6E6", "accent5": "FF3300", "accent6": "1A1A1A" }
```

**Signature:** Visible 12-column grid (hairline `line`s at 0.75 pt, `E6E6E6`) on content layouts, asymmetric placement (title in column 1–5, body in 7–12), one red geometric shape per slide (circle, square, or bar), titles flush-left at 48 pt with tight `charSpacing: -2`

---

### 12. Paper & Ink

**Vibe:** Editorial, literary, thoughtful

**Typography:** `Cormorant Garamond` + `Source Serif 4` · Office-safe fallback: `Garamond` · `Constantia`

**Theme:**
```json
{ "dk1": "1A1A1A", "lt1": "FAF9F7", "dk2": "FAF9F7", "lt2": "1A1A1A",
  "accent1": "C41E3A", "accent2": "5C5C5C", "accent3": "9A9A9A", "accent4": "E2DED6", "accent5": "8A6D3B", "accent6": "2B2B2B" }
```

**Background:** `make-background.py bg.png --solid faf9f7 --noise 0.025` (paper grain)

**Signature:** Drop caps (first letter as a separate 72 pt text box, crimson), pull quotes at 32 pt italic between two hairline rules, elegant horizontal rules (0.75 pt, `E2DED6`) instead of panels, generous 1 in margins

---

## Font Pairing Quick Reference

| Preset | Display Font | Body Font | Source | Office-safe fallback |
|--------|--------------|-----------|--------|----------------------|
| Bold Signal | Archivo Black | Space Grotesk | Google | Arial Black / Corbel |
| Electric Studio | Manrope | Manrope | Google | Gill Sans MT |
| Creative Voltage | Syne | Space Mono | Google | Impact / Consolas |
| Dark Botanical | Cormorant | IBM Plex Sans | Google | Perpetua / Candara |
| Notebook Tabs | Bodoni Moda | DM Sans | Google | Bodoni MT / Tw Cen MT |
| Pastel Geometry | Plus Jakarta Sans | Plus Jakarta Sans | Google | Century Gothic |
| Split Pastel | Outfit | Outfit | Google | Tw Cen MT |
| Vintage Editorial | Fraunces | Work Sans | Google | Bookman Old Style / Candara |
| Neon Cyber | Clash Display | Satoshi | Fontshare | Franklin Gothic Demi / Corbel |
| Terminal Green | JetBrains Mono | JetBrains Mono | JetBrains | Consolas |
| Swiss Modern | Archivo | Nunito | Google | Franklin Gothic Demi / Trebuchet MS |
| Paper & Ink | Cormorant Garamond | Source Serif 4 | Google | Garamond / Constantia |

Office-safe fallbacks are installed with Microsoft Office on both Windows and macOS. They are for machines where the distinctive pair cannot be installed; the distinctive pair is always the first choice.

---

## Type Scale by Density

| Role | Speaker-led | Reading-first |
|---|---|---|
| Title-slide title | 60–72 pt | 48–60 pt |
| Content-slide title | 40–48 pt | 32–40 pt |
| Section header | 28–32 pt | 22–26 pt |
| Body | 22–28 pt | 15–18 pt |
| Captions / chrome | 12–14 pt | 10–12 pt |
| Stat callout figure | 72–96 pt | 54–72 pt |

---

## DO NOT USE (Generic AI Patterns)

**Fonts:** Calibri, Arial, Aptos, Inter, Roboto, system fonts as display faces

**Colors:** `6366F1` (generic indigo), purple gradients on white, "corporate blue" title bars

**Layouts:** Everything centered, title + three equal bullet columns, identical card grids on every slide, accent lines under titles, header/footer color bars, sidebar stripes, edge-stripes on cards

**Decorations:** Realistic illustrations, stock-photo collages, gratuitous shadows and glows without purpose, WordArt effects

---

## Translating Web Design Values

The bold template pack and many design references are written for a 1920×1080 web stage. Convert with:

| Source | PowerPoint | Note |
|---|---|---|
| `px` (position/size) | `px ÷ 144` in | 1920 px = 13.333 in |
| `px` (font-size) | `px ÷ 2` pt | 112 px → 56 pt |
| `vw` | `× 0.1333` in or `× 9.6` pt | 7.5vw padding → 1.0 in |
| `vh` | `× 0.075` in or `× 5.4` pt | 5.5vh → 0.41 in |
| `em` letter-spacing | `em × font pt` → `charSpacing` | -0.02em at 56 pt → `charSpacing: -1.1` |
| unitless `line-height` | `lineSpacingMultiple` | 1.1 → `lineSpacingMultiple: 1.1` |
| `border-radius` | `rectRadius` (in) on `roundRect` | 24 px → 0.17 in |
| `box-shadow: 0 8px 32px rgba(0,0,0,.3)` | `shadow: { type: 'outer', blur: 8, offset: 2, angle: 90, color: '000000', opacity: 0.3 }` | blur/offset in pt, scaled down |
| `linear-gradient` / blur blobs / grid / noise | `make-background.py` PNG as background | not available as native fills |
| `::before` grid texture | `--grid` option | |
| CSS `opacity: 0.6` | `transparency: 40` | inverse |
