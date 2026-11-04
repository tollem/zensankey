# Options reference

Options are read from the `Sankey_Options` sheet (Excel, column A = key, column B = value) or
from the `Sankey_Options` text box next to the slide (PowerPoint, one `key = value` per line).
Keys are case-insensitive; unknown keys are ignored; empty values keep the default. The
**Options** button of the ribbon creates the sheet or the box with every key, its default and
a short description, so you rarely need to type keys by hand.

## Size and position

| Key | Default | Host | Description |
|---|---|---|---|
| `width` | 640 (Excel) / slide width − 72 (PowerPoint) | both | Diagram width in points (1 cm ≈ 28.35 pt) |
| `height` | 400 (Excel) / automatic (PowerPoint) | both | Diagram height in points |
| `left`, `top` | automatic | PowerPoint | Position of the diagram on the slide, in points |
| `margin` | 8 | both | Inner margin between the group bounds and the drawing |

In Excel the diagram is placed to the right of the selected range.

## Titles

| Key | Default | Description |
|---|---|---|
| `title` | | Bold title at the top left |
| `subtitle` | | Italic gray subtitle under the title |
| `column_titles` | | Section titles above each column, separated by `;`. Example: `Sources;Programs;Outcomes`. Leave an item empty to skip a column: `Sources;;Outcomes` |
| `column_title_size` | 9 | Font size of the section titles (upper case, gray) |

## Colors

| Key | Default | Description |
|---|---|---|
| `palette` | `Teal` | `Teal` (teal and warm accents), `Vivid` (saturated, colour-blind friendly order), `Pastel`, `Muted` (slate blues), `Ocean` (blues), `Custom` |
| `colors` | | Custom list used when `palette = Custom`: `#RRGGBB` codes separated by `;` |
| `color_by` | `node` | `node`: one palette color per node (column order, then appearance); `column`: all nodes of a column share a color |
| `node_color` | `palette` | `palette`, `gray` (dark gray nodes, a classic modern look with colored bands) or a fixed `#RRGGBB` |
| `link_color` | `gradient` | `gradient` (source color to target color), `source`, `target`, `gray` or a fixed `#RRGGBB`. The optional **Color** column of the table overrides this per band |
| `link_opacity` | 55 | Opacity of the bands in percent (100 = opaque) |
| `link_outline` | no | Thin white outline around each band (helps when bands overlap) |
| `text_color` | `#1C2B2D` | Color of labels and title |
| `background`, `background_color` | no, `#F7FBFB` | Rounded rectangle behind the whole diagram |

## Shapes

| Key | Default | Description |
|---|---|---|
| `link_shape` | `curved` | `curved` (Bezier bands) or `straight` (polygons) |
| `curvature` | 0.5 | 0.2 (tight) to 0.8 (very rounded); only for curved links |
| `node_width` | 14 | Width of the node bars in points |
| `node_padding` | 14 | Minimum vertical gap between nodes of the same column |
| `rounded_nodes` | no | Rounded corners on nodes |

## Labels and values

| Key | Default | Description |
|---|---|---|
| `labels` | `name_value` | `name_value`, `name`, `value` or `none` |
| `label_lines` | 2 | `2`: name on the first line, value in bold on the second; `1`: `Name (value)` |
| `outside_labels` | yes | Labels of the first column on the left of the nodes and labels of the last column on the right, outside the band area (margins are reserved automatically); `no`: labels next to the nodes inside the drawing |
| `label_background` | no | Translucent white box behind each label, useful when labels sit over bands |
| `link_values` | no | Writes the amount at the middle of every band tall enough to hold it |
| `value_format` | `standard` | `standard` (thousands separators), `k` (÷ 1 000), `M` (÷ 1 000 000), `percent` (share of the total entering the first column) |
| `decimals` | 0 | Number of decimals for amounts |
| `unit` | | Text appended to amounts, e.g. `CHF`, `GWh`, `t` |
| `font`, `font_size` | `Arial`, 10 | Font of all texts; the title uses 1.5 × `font_size` |

## Layout

| Key | Default | Description |
|---|---|---|
| `alignment` | `justify` | `justify`: nodes without outgoing flow are pushed to the last column and isolated sources sit next to their targets; `left`: columns follow the longest path from the sources only |
| `order` | `auto` | `auto`: nodes are ordered to reduce crossings (barycenter heuristic); `table`: order of first appearance in the table |

Forced columns: the optional 5th and 6th columns of the data table (Source column, Target
column) pin a node to a given column (0 = first). Useful for an external inflow that should
appear late in the diagram, or to align two diagrams.

## Recipes

Modern gradient look (default): `palette = Teal`, `link_color = gradient`, `link_opacity = 60`,
`outside_labels = yes`.

Monochrome report: `palette = Muted`, `node_color = gray`, `link_color = gray`,
`link_opacity = 45`, `label_background = yes`.

Financial statement with red costs and green profit: leave the palette as is, set
`node_color = gray`, and put `#E34948` or `#1BAF7A` in the **Color** column of each flow.

Presentation slide: `title = Where the money goes`, `subtitle = 2026 budget, CHF`,
`column_titles = Sources;Programs;Beneficiaries`, `font_size = 12`, `value_format = M`,
`decimals = 1`, `unit = M`, `background = yes`.

Compact dashboard tile: `labels = name`, `node_width = 8`, `node_padding = 6`, `font_size = 8`,
`margin = 2`.
