# ZenSankey

**Sankey diagrams in Excel and PowerPoint, drawn as native, editable shapes.**

ZenSankey is a pair of VBA add-ins (`.xlam` for Excel, `.ppam` for PowerPoint) that turn a
simple `Source | Target | Amount` table into a Sankey diagram made of real Office shapes:
curved bands, nodes, labels, titles. No picture, no web service, no data leaving your computer.
The result is a grouped shape you can resize, recolor and retouch like any other object, and
copy from Excel to PowerPoint without losing anything.

![Example: income statement rendered by the ZenSankey layout engine](assets/example-income-statement.png)

## Features

- Works in **Excel** (from a selected range) and **PowerPoint** (from a table or a text box).
- **Native shapes**: every band, node and label is editable; the whole diagram is one group.
- Automatic layout: columns from the flow structure, sinks pushed to the right, node ordering
  that reduces crossings, proportional heights, optional forced columns.
- **Styling options**: five palettes plus custom colors, color by node or by column,
  gradient / source / target / gray / fixed link colors, opacity, outline, curved or straight
  links, node width and padding, rounded nodes.
- **Text**: font, size and color, labels on one or two lines, labels outside the diagram for
  the first and last columns, translucent label backgrounds, values on the bands, value formats
  (standard, thousands, millions, percent), decimals and unit.
- **Titles**: title, subtitle and section titles above each column; optional background.
- Tolerant number parsing (`3 000`, `3'000`, `2 650,5`, `1,234.5`).
- Windows and macOS, Office 2016 and later (Microsoft 365 included). No installation of anything
  else, no administrator rights.

## Installation

Download `ZenSankey.xlam` and `ZenSankey.ppam` from the [`dist`](dist) folder or from the
[Releases](../../releases) page, then follow one of the methods below.

### Windows, scripted

1. Close Excel and PowerPoint.
2. Open a PowerShell window in the folder that contains the add-ins and run:

   ```powershell
   powershell -ExecutionPolicy Bypass -File install\install-windows.ps1
   ```

   The script copies the add-ins to `%APPDATA%\Microsoft\AddIns`, unblocks them and registers
   them for Excel and PowerPoint. Use `-Uninstall` to remove them.

### Windows, manual

1. Right-click each downloaded file > **Properties** > tick **Unblock** > OK.
2. Move the files to a permanent folder (for example `%APPDATA%\Microsoft\AddIns`).
3. Excel: **File > Options > Add-ins**, at the bottom choose **Excel Add-ins** > **Go…** >
   **Browse…**, select `ZenSankey.xlam`, tick it.
4. PowerPoint: **File > Options > Add-ins**, choose **PowerPoint Add-ins** > **Go…** >
   **Add New…**, select `ZenSankey.ppam`, accept the macro prompt.

### macOS

1. Run `bash install/install-macos.sh` (copies the files to the Office add-ins folder and
   removes the quarantine flag), or copy them yourself to
   `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized/`.
2. Excel: **Tools > Excel Add-ins…**, tick **ZenSankey** (or **Browse…** to the file).
3. PowerPoint: **Tools > PowerPoint Add-ins…**, click **+**, choose `ZenSankey.ppam`.

### Try it without installing

The `dist` folder also contains `ZenSankey-test.xlsm` and `ZenSankey-test.pptm`: a
workbook and a presentation that embed the same macros and ribbon. Open one, enable macros,
and use the **Sankey** tab.

### If macros are blocked

Organisations often restrict macros. If the Sankey tab does not appear or the buttons do
nothing, add the folder that contains the add-ins to the **Trusted Locations**
(File > Options > Trust Center > Trust Center Settings > Trusted Locations), or ask your IT
department. The add-ins contain no external calls and no network access; the source code is in
[`src`](src).

## Usage

### Excel

1. Build a table with one row per flow: **Source | Target | Amount**. A header row is optional.
   Three optional columns can follow: **Color** (`#RRGGBB`, forces the color of that band),
   **Source column** and **Target column** (0 = first column; forces the horizontal position
   of a node).
2. Select the table and click **Sankey > Draw Sankey**. The diagram appears to the right of the
   selection as a group named `Sankey hhmmss`.
3. Click **Sankey > Options** once to add a `Sankey_Options` sheet to the workbook. Edit the
   values in column B (drop-down lists where relevant) and draw again. `width` and `height`
   set the size of the diagram in points.
4. **Example** inserts a sample table (income statement) at the active cell.

### PowerPoint

1. Insert a table with the same columns (**Insert > Table**, or paste cells copied from Excel),
   or a text box with one flow per line in the form `Source [amount] Target`
   (`Source ; Target ; amount` and tab-separated lines also work).
2. Select the table or the text box and click **Sankey > Draw Sankey**. The diagram is placed
   below the table when there is room, otherwise in the content area of the slide.
3. Click **Sankey > Options** to add an options box named `Sankey_Options` next to the slide
   (outside the slide, so it is not shown during the slide show). Edit the lines
   `key = value`; `left`, `top`, `width`, `height` set the position and size.

### Updating a diagram

Draw again from the (modified) table and delete the previous group. Because the result is a
plain group of shapes, you can also edit it directly: move labels, change a color, enlarge
the group.

## Options

All options have sensible defaults; set only what you need. The full reference with examples
is in [`docs/options.md`](docs/options.md).

| Key | Default | What it does |
|---|---|---|
| `title`, `subtitle` | | Texts above the diagram |
| `column_titles` | | Section titles above each column, separated by `;` |
| `palette` | `Teal` | `Teal`, `Vivid`, `Pastel`, `Muted`, `Ocean`, `Custom` |
| `colors` | | Custom colors `#RRGGBB;#RRGGBB;…` (with `palette = Custom`) |
| `color_by` | `node` | Palette assignment: `node` or `column` |
| `node_color` | `palette` | `palette`, `gray` or `#RRGGBB` |
| `link_color` | `gradient` | `gradient`, `source`, `target`, `gray` or `#RRGGBB` |
| `link_shape`, `curvature` | `curved`, `0.5` | Band shape |
| `link_opacity`, `link_outline` | `55`, `no` | Band opacity in %, white outline |
| `node_width`, `node_padding`, `rounded_nodes` | `14`, `14`, `no` | Node geometry |
| `font`, `font_size`, `text_color` | `Arial`, `10`, `#1C2B2D` | Text style |
| `labels`, `label_lines` | `name_value`, `2` | Label content and layout |
| `outside_labels`, `label_background` | `yes`, `no` | Label placement and readability |
| `link_values` | `no` | Amount written on bands that are tall enough |
| `value_format`, `decimals`, `unit` | `standard`, `0`, | `standard`, `k`, `M`, `percent` |
| `background`, `background_color` | `no`, `#F7FBFB` | Rounded background |
| `alignment`, `order`, `margin` | `justify`, `auto`, `8` | Layout rules |

Booleans accept `yes`/`no`, `true`/`false`, `1`/`0` (and `oui`/`non`).

## Building from source

The add-ins are built from the `.bas` files in `src/` with LibreOffice as the VBA compiler
and a Python packaging step. See [`docs/building.md`](docs/building.md). To contribute, read
[`CONTRIBUTING.md`](CONTRIBUTING.md).

## Troubleshooting

See [`docs/troubleshooting.md`](docs/troubleshooting.md) for the common cases: tab missing,
macros disabled, error messages, accents, large diagrams.

## License and attribution

ZenSankey is released under the [Apache License 2.0](LICENSE).
Copyright (c) 2026 **Samuel Mellot**, original author. The [`NOTICE`](NOTICE) file carries the
attribution and must be kept in redistributions and derivative works, including forks
(Apache License 2.0, Section 4). See also [`AUTHORS.md`](AUTHORS.md) and
[`CITATION.cff`](CITATION.cff).
