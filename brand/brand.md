# xlsx-stream-workbook — brand system

## Strategy

| | |
|---|---|
| **Category** | Developer tool · data export infrastructure |
| **Audience** | Backend and data engineers who export to Excel at volume and have already hit a memory wall |
| **Personality** | Precise, restrained, unpretentious, engineering-credible |
| **Emotional promise** | Your data will not take the process down |
| **Cultural position** | The boring, correct tool you trust in production — not a startup landing page |
| **Visual world** | Blueprint, not brochure. Hairlines, one signal colour, generous emptiness |
| **What to avoid** | The green spreadsheet icon, cloud-and-bolt developer clichés, purple-blue AI glow, badge soup |

## Core metaphor

**Data passes through a document.**

Every Excel writer on npm holds the whole workbook and writes it at the end. This one does not: rows are streamed into a temporary file as each sheet is added, and those files are merged into one package on save. The distinctive truth of the library is that it is *permeable* — the work goes through, and nothing accumulates.

## Symbol logic

Methods combined: **Product Action** (merge many into one) and **Construction Geometry** (a measured frame).

**The mark is a page with a bar passing through it.**

- The square frame is the sheet — a document, measured, not decorative.
- The bar is the row stream. It enters from the left, crosses the page, and exits on the right.
- Where the bar crosses, it **replaces** the frame's edge. The document is opened by the data passing through it. That knockout is the whole idea in one detail, and it is what keeps the mark from being "a square with a line in it".

It is deliberately **not** a grid of cells. Every competitor's mark is a grid; a grid would make this one forgettable. One frame and one bar is also the only version that survives at 16px, which is the only size that decides whether a mark is real.

Construction: a 64-unit grid, frame inset 10, frame stroke 2.5, bar 8 units tall on the vertical centre, extending exactly 8 units past the frame on each side. The bar's height and the frame's gap are the same value so the replacement is seamless.

## Palette

Dark Developer / Builder mode. One accent, and it carries the entire system.

| Token | Value | Role |
|---|---|---|
| `ink` | `#111318` | Type, frame, dark canvas |
| `paper` | `#f7f6f3` | Light canvas, knockouts |
| `accent` | `#c2410c` | The bar, focal marks, one chip per surface |
| `muted` | `#6b7280` | Secondary type, captions |
| `rule` | `rgba(17,19,24,0.12)` | Hairlines, gutters |

The accent is a deep burnt orange rather than the tangerine used by the diagrams, so the brand reads as its own thing and not as a skin of the documentation. It is darker and less decorative on purpose: this is a tool, not a launch.

## Typography

- **Wordmark** — one weight, sentence case, `ink`. Never letterspaced, never italicised, never in a display face.
- **Interface type** — system sans at 14/11px for panels.
- **Technical type** — monospace at 10px, `muted`, for labels, file names, and commands only.
- No third font. A three-font system on a nine-panel board is two fonts too many.

## Tagline

> Rows pass through. Nothing accumulates.

Specific to what the library does, falsifiable, and it names the trade-off rather than claiming a benefit.

## Applications

| Surface | The mark appears as |
|---|---|
| README / npm | Full lockup, light, above the first heading |
| GitHub social preview | Mark + wordmark on `ink`, tagline in `accent` |
| favicon / repo icon | Mark only, 16px, on `paper` |
| Terminal | The install line, monospace, no mark |
| Diagram pages | Not used — the diagrams carry their own palette and must not be rebranded |

That last row is a real constraint, not an oversight. The three diagrams in `docs/diagrams/` already have a skin, and a brand system that repaints them would be a brand system arguing with itself.
