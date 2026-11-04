# Style keys reference

The ribbon is the normal way to set the style (see the README). Internally every control maps
to a `key = value` pair; the full set is stored in the alternative text of each diagram group,
and can be set from VBA with `SkSetOption "key", "value"` or `SkApplyOptionText` (one pair per
line) before calling `SkDraw`. Keys are case-insensitive; unknown keys are ignored; empty
values keep the default.

## Size

In Excel the diagram starts at 620 x 360 points to the right of the selected range and grows as
needed so that every label fits. In PowerPoint it fits the slide (below the source table when
there is room, otherwise the content area) and the text is reduced when needed. Resize the
group and click **Update** to redraw at another size.

| Key | Default | Description |
|---|---|---|
| `margin` | 8 | Inner margin between the group bounds and the drawing |

## Titles

| Key | Default | Description |
|---|---|---|
| `title` | | Bold title at the top left |
| `subtitle` | | Italic gray subtitle under the title |
| `column_titles` | | Section titles above each column, separated by `;`. Example: `Sources;Programs;Outcomes`. Leave an item empty to skip a column: `Sources;;Outcomes` |
| `column_title_size` | 10 | Font size of the section titles (bold, gray) |

## Colors

| Key | Default | Description |
|---|---|---|
| `palette` | `Teal` | `Teal` (teal, violet, red, green, orange...), `Vivid` (saturated), `Pastel`, `Muted` (slate blues), `Ocean` (blues), `Warm`, `Mono` (shades of teal), `Custom` |
| `colors` | | Custom list used when `palette = Custom`: `#RRGGBB` codes separated by `;` |
| `color_by` | `flow` | `flow`: a node keeps the color of the source that sends it the most, when that is at least half of the source's output; other nodes get the next palette color (so the main flow keeps one color and branches get accents); `node`: one palette color per node (column order, then appearance); `column`: all nodes of a column share a color |
| `node_color` | `palette` | `palette`, `gray` (dark gray nodes, a classic modern look with colored bands) or a fixed `#RRGGBB` |
| `link_color` | `target` | `target` (color of the destination node), `source`, `gradient` (source color to target color), `gray` or a fixed `#RRGGBB`. The optional **Color** column of the table overrides this per band and also colors the target node |
| `link_opacity` | 45 | Opacity of the bands in percent (100 = opaque) |
| `link_outline` | no | Thin white outline around each band (helps when bands overlap) |
| `text_color` | `#1C2B2D` | Color of labels and title |
| `background`, `background_color` | no, `#F7FBFB` | Rounded rectangle behind the whole diagram |

## Shapes

| Key | Default | Description |
|---|---|---|
| `link_shape` | `curved` | `curved` (Bezier bands) or `straight` (polygons) |
| `curvature` | 0.5 | 0.2 (tight) to 0.8 (very rounded); only for curved links |
| `node_width` | 12 | Width of the node bars in points |
| `node_padding` | 16 | Minimum vertical gap between nodes of the same column (raised automatically to the label height) |
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

Default look: `palette = Teal`, `color_by = flow`, `link_color = target`, `link_opacity = 45`,
`outside_labels = yes`.

Monochrome report: `palette = Muted`, `node_color = gray`, `link_color = gray`,
`link_opacity = 45`, `label_background = yes`.

Financial statement with red costs and green profit: put `#E34948` or `#1BAF7A` in the
**Color** column of each flow (the band and its target node take that color).

Presentation slide: `title = Where the money goes`, `subtitle = 2026 budget, CHF`,
`column_titles = Sources;Programs;Beneficiaries`, `font_size = 12`, `value_format = M`,
`decimals = 1`, `unit = M`, `background = yes`.

Compact dashboard tile: `labels = name`, `node_width = 8`, `font_size = 8`, `margin = 2`.
