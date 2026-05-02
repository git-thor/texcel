# Changelog

All notable changes to this project will be documented in this file.

## [0.1.0] - 2026-05-02

### Added

- Initial release of Excel ↔ LaTeX table converter
- Bidirectional conversion:
  - Excel to LaTeX (`excel_to_latex`)
  - LaTeX to Excel (`latex_to_excel`)
- Excel to LaTeX features:
  - Math notation preservation (`$\gamma=0$`, `$\pi_H(s)$`, etc.)
  - Bold/italic formatting from Excel cells
  - Multiline output format for readability
  - Automatic column header comments
  - Smart escaping of special LaTeX characters
  - Empty row/column handling (midrule and separators)
  - Configurable tabularx/tabular output
  - Landscape/portrait mode support
  - Custom first column width
  - Font size options (tiny, scriptsize, footnotesize, small, normalsize)
  - Table captions and labels
- LaTeX to Excel features:
  - Parse `tabular` and `tabularx` environments
  - Extract cell content with formatting markers
  - Handle multi-row tables with midrules
  - Preserve bold text formatting (`\textbf{}`)
- Column handling:
  - Leading/trailing empty columns preserved as regular columns
  - Middle empty columns converted to `|` separators
- Command-line interface with argparse:
  - Input/output file arguments
  - Sheet selection (`--sheet`)
  - List sheets (`--list-sheets`)
  - Verbose output with dimensions and formatting stats
  - Helpful usage examples in `--help`
- CLI entry points:
  - `excel-to-latex`
  - `latex-to-excel`
- Comprehensive test suite (29 tests):
  - Roundtrip conversion tests
  - Empty column handling tests
  - Math symbol preservation tests
  - Special character escaping tests
  - Formatting preservation tests
- README.md with usage examples and installation instructions
- MIT License
- Type hints and Google-style docstrings throughout codebase

### Fixed

- URL-decoding of compressed Draw.io payloads after zlib decompression
- Edge path newline formatting (literal `\n` strings → actual newlines)
- Ipe arrow attribute syntax (separate `arrow` and `rarrow` for end/start heads)
- Group container coordinate offset accumulation with recursive transforms
- Invalid edge point handling (partial `None` coordinates resolved to group boundaries)
- Prevention of group container rendering as shapes
- Leading empty columns now preserved correctly in LaTeX output
- Multiline format properly handles consecutive empty cells
- Empty middle columns correctly create `|` separators in column spec

### Technical

- Python 3.13+ compatibility
- Package metadata in `pyproject.toml`
- Dev dependencies: pytest, ruff, mypy
- Dependencies: pandas, openpyxl, latex2python
