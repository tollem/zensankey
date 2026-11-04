# Building from source

The add-ins are generated from three VBA modules in `src/`:

| Module | Role |
|---|---|
| `SankeyCore.bas` | Layout engine and drawing, shared by both hosts (uses only the Office `Shapes` object model) |
| `SankeyExcel.bas` | Excel entry points: selection reading, `Sankey_Options` sheet, example, ribbon callbacks |
| `SankeyPowerPoint.bas` | PowerPoint entry points: table or text reading, `Sankey_Options` box, example, ribbon callbacks |

## How the binary project is produced

Office stores macros in a binary `vbaProject.bin` (MS-OVBA format). Instead of writing that
format by hand, the build uses **LibreOffice** as a compiler: `build/lo_vba_export.py` starts a
headless LibreOffice, loads the `.bas` files into a VBA-compatible Basic library and saves a
`.xlsm`, from which `vbaProject.bin` is extracted. The project is host-agnostic, so the same
mechanism produces the Excel and the PowerPoint projects.

`build/package.py` then wraps each project into a package: a minimal workbook for the `.xlam`
(content type `application/vnd.ms-excel.addin.macroEnabled.main+xml`), a minimal presentation
for the `.ppam` (`application/vnd.ms-powerpoint.addin.macroEnabled.main+xml`), the relationship
to `vbaProject.bin` and the ribbon definition `customUI/customUI14.xml`. Macro-enabled test
documents (`.xlsm`, `.pptm`) are produced the same way.

## Requirements

- LibreOffice 7.x or later (`soffice` on the PATH)
- Python 3.10+ with the `uno` module (package `python3-uno` on Debian/Ubuntu; the Python
  bundled with LibreOffice also works), `openpyxl` and `python-pptx`
- Optional for inspection: `oletools` (`olevba dist/ZenSankey.xlam` lists the macros)

## Commands

```bash
python3 build/test_layout.py   # compiles the modules in LibreOffice and runs tests/SkTest.bas
python3 build/build.py         # builds dist/ZenSankey.xlam, .ppam and the test documents
```

## Rules for the sources

- **Pure ASCII.** Office on macOS reads module text with the Mac Roman code page, so any
  accented character in the source is corrupted there. Generate such characters at run time
  with `ChrW()`. `build/build.py` refuses non-ASCII sources.
- **No VBA reserved words as identifiers** (`Scale`, `Line`, `Name`, `Width`, `Text`, `Date`,
  `Seek`, `Tab`, ...). LibreOffice Basic accepts some of them, Office does not.
- Keep the drawing code late-bound (`As Object`): the engine must work in both hosts and
  needs no reference beyond the default ones (`stdole`, `Office`).
- LibreOffice can compile and run the layout engine but does not implement `BuildFreeform`
  or `AddTextbox`: the drawing path must be tested in Excel and PowerPoint.

## Editing without rebuilding

For a quick change, open the `.xlam` or `.ppam` in Office, press Alt+F11 (Option+F11 on
macOS), edit the modules and save. To replace the ribbon, edit `customUI/customUI14.xml`
inside the package (it is a ZIP archive) or rebuild.
