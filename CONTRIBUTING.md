# Contributing

Thank you for considering a contribution.

- **Bugs and ideas**: open an issue with your Office version (Windows or macOS), the host
  (Excel or PowerPoint), a minimal flow table and, for errors, the exact message and the
  highlighted line in the VBA editor (Alt+F11).
- **Code**: the add-ins are built from the `.bas` sources in `src/`. Keep the sources pure
  ASCII (use `ChrW()` for other characters) and avoid VBA reserved words as identifiers.
  Run `python3 build/test_layout.py` (needs LibreOffice) before opening a pull request, then
  test in Excel and PowerPoint. See `docs/building.md`.
- **License**: by contributing you agree that your contribution is licensed under the
  Apache License 2.0. The `NOTICE` file and the original author attribution must be kept.
