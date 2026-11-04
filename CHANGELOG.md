# Changelog

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
