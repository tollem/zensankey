# Troubleshooting

**The Sankey tab does not appear.**
The add-in is not loaded. Excel: File > Options > Add-ins > Excel Add-ins > Go, tick
ZenSankey (Browse to the file if needed). PowerPoint: the same dialog with PowerPoint
Add-ins > Add New. On macOS use Tools > Excel Add-ins / PowerPoint Add-ins. If the add-in is
ticked but the tab is still missing, macros are probably disabled by policy: add the folder of
the add-in to the Trusted Locations, or ask your IT department.

**"Microsoft has blocked macros from running because the source of this file is untrusted."**
The file carries the "downloaded from the Internet" mark. Right-click the file > Properties >
Unblock, or run `install\install-windows.ps1`, which does it for you.

**The buttons do nothing / "Cannot run the macro".**
Macro security is set to "Disable all macros without notification". Choose "Disable with
notification" or use a Trusted Location.

**"Select the range that contains the flows first".**
Excel needs a selected range with at least three columns (Source, Target, Amount).
PowerPoint needs a selected table or text box.

**"No valid flow found".**
Amounts must be positive numbers; empty rows are skipped. In PowerPoint text boxes, use one
line per flow: `Source [amount] Target`.

**Compile error in a hidden module.**
Report the exact message and, if possible, the highlighted line (Alt+F11 opens the editor).
Known causes are fixed in version 1.1 and later: reserved identifiers and non-ASCII sources.

**Accented characters look wrong on macOS.**
Update to version 1.1 or later; sources are pure ASCII and accents are generated at run time.

**Labels overlap.**
Labels are measured on the real text, so this only happens in very dense diagrams or when a
diagram was resized much smaller. Enlarge the group and click **Update**, choose a smaller
label size, `Name only` labels, or enable **Label background**.

**The diagram is too small or in the wrong place.**
Move or resize the group by hand, then click **Update**: the diagram is redrawn at the new
size (in PowerPoint the text is reduced to fit when needed).

**A ribbon change did nothing.**
Style controls act on the *selected* diagram (click on it first) and on the next diagrams
drawn. Text boxes (Title, Subtitle, Sections, Unit) apply when you press Enter.

**Update says the source table was not found.**
The diagram remembers the sheet and range (Excel) or the slide shape (PowerPoint) it was drawn
from. If the table was moved, deleted or renamed, select the table and click **Draw Sankey**
again.

**A node is in the wrong column.**
Use the Source column / Target column data columns (0 = first column) to pin it, or set
`alignment = left` to disable the justification of sinks.

**Gradient links are not visible.**
Some older Office versions do not support gradient stops through VBA; the add-in then falls
back to a solid fill in the source color. Set `link_color = source` to make it explicit.
