# Changelog

## 2.1.0 - 2026-11-04

- New default look: flat translucent bands colored by target, node colors that follow the main
  flow (`color_by = flow`), balanced column spacing, labels measured on the real text and
  wrapped when long, section titles in bold gray (no longer upper case).
- Style is set from the ribbon (palette, colors, links, opacity, shape, nodes, labels, font,
  number format, unit, titles, background, order) and applied immediately to the selected
  diagram; **Reset style** restores the defaults. The `Sankey_Options` sheet and text box are
  gone; the style is stored with each diagram.
- **Update** redraws the selected diagram from its source table, at its current position and
  size. In Excel the diagram grows so that every label fits; in PowerPoint the text is reduced
  to fit the slide.
- A colored flow (Color column) also colors its target node. New palettes `Warm` and `Mono`.
- **Example** draws the sample table directly.

## 2.0.0 - 2026-11-04

- Options system: palettes (Teal, Vivid, Pastel, Muted, Ocean, Custom), node and link colors,
  gradient links, straight or curved links, link outline, rounded nodes, fonts and text color,
  label content and placement, label background, link values, value formats (standard, k, M,
  percent), unit, title, subtitle, section titles above columns, background, alignment, order.
- Excel: `Options` button creates a `Sankey_Options` sheet with every option, defaults and
  drop-down choices; `width` and `height` control the diagram size.
- PowerPoint: `Options` button adds a `Sankey_Options` text box next to the slide; `left`,
  `top`, `width`, `height` control the placement.
- Outside labels for the first and last columns, with automatic margins.
- English user interface and documentation.

## 1.1.0 - 2026-11-04

- Pure ASCII sources (accented characters generated at run time) for Office on macOS.
- Fixed a compile error caused by a reserved identifier.

## 1.0.0 - 2026-11-04

- First release: Excel (.xlam) and PowerPoint (.ppam) add-ins drawing Sankey diagrams as
  native, editable shapes from a Source | Target | Amount table.
