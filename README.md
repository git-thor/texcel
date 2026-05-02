# Texcel - Excel to LaTeX Table Converter

A best-effort converter for transforming Excel tables into LaTeX format, enabling seamless integration of tabular data into LaTeX documents.

## Why Convert Excel to LaTeX?

While Excel is excellent for data management and calculations, **LaTeX** offers crucial advantages for academic and technical publishing:

### Full LaTeX Integration

LaTeX tables provide:

-   **Perfect font matching**: Your table fonts exactly match your document's font family, size, and style
-   **Professional typesetting**: Beautiful mathematical formulas, consistent spacing, and professional layout
-   **Native integration**: Use `\cref{}` for cross-references, `\ac{}` for acronyms, and all LaTeX commands
-   **Publication quality**: High-quality output suitable for journals, theses, and conferences

### Advanced Features

-   **Math mode support**: Preserve `$\gamma=0$`, `$\pi_H(s)$`, and other mathematical notation
-   **Bold/italic formatting**: Maintain cell formatting from Excel
-   **Multiline format**: Each column on a new line for better readability in source code
-   **Column comments**: Automatic header comments show which column data belongs to
-   **Smart escaping**: Proper handling of `%`, `&`, `$`, and other special characters

## Installation

### Prerequisites

1.  **Python 3.13+** with `pip` or `uv`
2.  **TeX Live** or similar LaTeX distribution (for compiling the output)

### Install Dependencies

Using `uv` (recommended):

```bash
uv sync
```

Or using standard pip:

```bash
pip install .
```

For development purposes, install as editable:

```bash
pip install -e .
```

## Usage

### Basic Conversion

#### Excel to LaTeX

```bash
# Using uv
uv run excel2latex.py input.xlsx -o output.tex

# Using Python directly
python excel2latex.py input.xlsx -o output.tex

# Using installed command
texcel input.xlsx -o output.tex
```

#### LaTeX to Excel

```bash
# Using uv
uv run latex2excel.py input.tex -o output.xlsx

# Using installed command
latex2excel input.tex -o output.xlsx
```

### Options

#### Excel to LaTeX

```bash
uv run excel2latex.py --help
```

Options:
```
  -o, --output FILE         Output .tex file path (default: output_table.tex)
  -s, --sheet NAME          Sheet name or index (default: 0)
  --no-multiline            Use compact single-line format (default is multiline)
  --no-escape               Don't escape special characters (preserve math notation)
  --no-formatting           Don't preserve bold formatting from Excel
  -c, --caption TEXT        Table caption
  --label TEXT              LaTeX label (e.g., tab:mytable)
  --skip-rows N             Skip N rows at beginning
  --landscape               Use landscape mode (default: True)
  --no-landscape            Use portrait mode
  -w, --width TEXT          First column width (default: 2cm)
  -f, --font SIZE           Font size (tiny, scriptsize, footnotesize, small, normalsize)
```

#### LaTeX to Excel

```bash
uv run latex2excel.py --help
```

Options:
```
  -o, --output FILE         Output .xlsx file path (default: output.xlsx)
  -s, --sheet NAME          Sheet name (default: Sheet1)
  --keep-empty              Keep empty rows in output
  --no-formatting           Don't preserve bold formatting
  -t, --table-index N       Table index to extract (0-indexed, default: 0)
```

### Examples

```bash
# Standard Excel to LaTeX conversion (multiline format with comments)
uv run excel2latex.py table.xlsx -o table.tex

# Convert with caption and label
uv run excel2latex.py table.xlsx -o table.tex -c "My Table" -l "tab:mytable"

# Compact single-line format
uv run excel2latex.py table.xlsx -o table.tex --no-multiline

# Preserve math notation without escaping
uv run excel2latex.py table.xlsx -o table.tex --no-escape

# Custom first column width and font size
uv run excel2latex.py table.xlsx -o table.tex -w "3cm" -f footnotesize

# Convert LaTeX back to Excel
uv run latex2excel.py table.tex -o table.xlsx

# Extract specific table from multi-table LaTeX file
uv run latex2excel.py tables.tex -o table.xlsx -t 1
```

## Supported Features

The converter handles the most common Excel/LaTeX table features needed for LaTeX workflows:

### Excel → LaTeX

-   **Cell content**: Text, numbers, dates, formulas
-   **Math notation**: `$\gamma=0$`, `$\pi_H(s)$`, `$\Delta[k_p, k_i]$` preserved
-   **Bold formatting**: Bold cells wrapped in `\textbf{}`
-   **Empty rows**: Converted to `\midrule` for visual separation
-   **Empty columns**: Converted to `|` separators in column specification
-   **Special characters**: Automatic escaping of `%`, `&`, `$`, `#`, `_`, etc.
-   **Column comments**: Header names added as comments for readability

### LaTeX → Excel

-   **Tabular/tabularx**: Both standard and tabularx environments
-   **Bold formatting**: `\textbf{}` cells marked as bold in Excel
-   **Math notation**: Preserved as LaTeX text (not converted to Unicode)
-   **Comments stripped**: LaTeX comments (`%`) properly removed
-   **Empty cells**: Preserved as empty Excel cells
-   **Multiple tables**: Extract specific tables by index

### Limitations

-   Complex Excel formatting (colors, merged cells) not preserved
-   Images and charts not supported
-   Excel formulas converted to values only
-   LaTeX commands beyond basic formatting are stripped

## Workflow

### Excel to LaTeX Workflow

1.  **Create your table** in Excel with data and formatting
2.  **Add math notation** using LaTeX syntax (e.g., `$\gamma=0$`)
3.  **Run conversion**: `texcel table.xlsx -o table.tex`
4.  **Include in LaTeX**: `\input{table.tex}` or copy the content
5.  **Compile**: Your table appears with perfect formatting

### LaTeX to Excel Workflow

1.  **Export LaTeX table** from your document
2.  **Run conversion**: `latex2excel table.tex -o table.xlsx`
3.  **Edit in Excel**: Modify data, add rows, etc.
4.  **Convert back**: `texcel table.xlsx -o table_updated.tex`
5.  **Replace in document**: Update your LaTeX file

## Example Output

### Input Excel
| | Tietze | Kampmeier | McC/Law |
|---|---|---|---|
| **Learning Paradigm** | Batch, model-based | Sequential, PPO | Meta-RL |
| **Math Test** | `$\gamma=0$` | `$\pi_H(s)$` | `100%` |

### Output LaTeX (Multiline Format)
```latex
\begin{table}[p]
    \centering
    \begin{tabularx}{\linewidth}{@{}>{\raggedright\arraybackslash}p{2cm} *{3}{>{\centering\arraybackslash}X}@{}}
        \toprule
             &    \textbf{Tietze}  % Tietze
             &    \textbf{Kampmeier}  % Kampmeier
             &    \textbf{McC/Law}  % McC/Law
             \\

            \textbf{Learning Paradigm}
             &    Batch, model-based  % Tietze
             &    Sequential, PPO  % Kampmeier
             &    Meta-RL  % McC/Law
             \\

            \textbf{Math Test}
             &    $\gamma=0$  % Tietze
             &    $\pi_H(s)$  % Kampmeier
             &    100\%  % McC/Law
             \\
        \bottomrule
    \end{tabularx}
\end{table}
```

## Technical Structure

```
excel2latex/
├── excel_to_latex.py     # Excel to LaTeX conversion
├── latex_to_excel.py     # LaTeX to Excel conversion
├── pyproject.toml        # Project configuration
├── test_comprehensive.py # Comprehensive test suite
├── test_roundtrip.py     # Round-trip tests
└── README.md             # This file
```

## Testing

Run the test suite with pytest:

```bash
uv run pytest
```

All 29 tests cover:
- Empty column handling
- LaTeX comment stripping
- Empty rows as midrule
- Math symbol preservation
- Percent sign escaping
- Bold formatting
- Round-trip conversion

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests for:

-   Support for additional Excel features (colors, merged cells)
-   Better formatting preservation
-   Performance improvements
-   Additional output formats

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

-   The LaTeX community for excellent documentation and packages
-   OpenPyXL for Excel file handling
-   Pandas for data manipulation
-   Academic publishing community for feedback and testing
