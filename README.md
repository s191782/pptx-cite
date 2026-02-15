# PPTX Citation Manager

A Python tool to manage LaTeX-style citations in PowerPoint presentations using BibTeX/BibLaTeX references.

## Features

- **LaTeX-style Citations**: Use familiar `\cite{key}` commands in your PowerPoint slides
- **Multiple Citation Styles**: Support for numeric, author-year, and author-only citations
- **Automatic Bibliography**: Generates a references slide where you place the `\bibliography` command
- **BibTeX/BibLaTeX Support**: Works with standard .bib files, supports 16+ entry types
- **Format Preservation**: Maintains text formatting, bullets, numbering, and styling
- **Bibliography Overflow**: Automatically creates continuation slides if bibliography doesn't fit
- **Smart Author Formatting**: Handles complex author names with curly brackets, uses "et al." for >3 authors
- **Citation Order Preservation**: Citations numbered in order of first appearance, not alphabetically
- **Broad Shape Support**: Processes text boxes, shapes, grouped shapes, and tables
- **Optional .bib File**: Defaults to presentation name with .bib extension
- **Configurable**: Customize citation and bibliography commands via YAML config

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Process your presentation (uses presentation.bib automatically)
python pptx_cite.py presentation.pptx

# Or specify a different .bib file
python pptx_cite.py presentation.pptx references.bib
```

## Usage

### Basic Usage

```bash
# Uses presentation.bib automatically
python pptx_cite.py presentation.pptx

# Or specify a different .bib file
python pptx_cite.py presentation.pptx references.bib
```

**Output**: Creates `presentation_cited.pptx` with formatted citations and bibliography.

### Advanced Options

```bash
# Custom configuration
python pptx_cite.py presentation.pptx -c config.yaml

# Specify output filename
python pptx_cite.py presentation.pptx -o final_presentation.pptx

# Custom starting number for citations (default: 1)
python pptx_cite.py presentation.pptx -s 5  # Citations start at [5], [6], [7]...

# Combine options
python pptx_cite.py presentation.pptx references.bib -c config.yaml -o output.pptx -s 10
```

## How to Use in PowerPoint

1. **Add Citations**: In your slides, type citations like:
   ```
   This is a fact \cite{smith2020}.
   Multiple citations work too \cite{jones2019,doe2021}.
   ```

2. **Create Bibliography Slide**: Add a slide where you want the bibliography and type:
   ```
   \bibliography
   ```

3. **Run the Script**: Process your presentation with the script

4. **Result**: Citations will be replaced with formatted references and the bibliography slide will be populated

## Configuration

Create a `config.yaml` file to customize behavior:

```yaml
# Citation command to search for (default: \cite)
cite_command: '\cite'

# Bibliography command (default: \bibliography)
bibliography_command: '\bibliography'

# Citation style: numeric, numeric-parens, author-year, or author
cite_style: 'numeric'
```

### Citation Styles

- **numeric** (default): `[1]`, `[2, 3]`
- **numeric-parens**: `(1)`, `(2, 3)`
- **author-year**: `(Smith, 2020)`, `(Jones et al., 2019)`
- **author**: `(Smith)`, `(Jones et al.)`

### Custom Starting Number

For numeric citation styles, you can customize the starting number using the `-s` or `--start-number` option:

```bash
python pptx_cite.py presentation.pptx -s 5  # Citations: [5], [6], [7], etc.
python pptx_cite.py presentation.pptx --start-number 10  # Citations: [10], [11], [12], etc.
```

This is useful when:
- Continuing citation numbering from a previous document
- Creating supplementary materials that extend the main document's references
- Building multi-part presentations with consistent numbering across parts

## Supported BibTeX/BibLaTeX Formats

### Entry Types (16+)
✅ Articles & Books: `@article`, `@book`, `@inproceedings`, `@incollection`, `@inbook`  
✅ Theses: `@mastersthesis`, `@phdthesis`, `@thesis` (BibLaTeX)  
✅ Reports & Documentation: `@techreport`, `@manual`, `@proceedings`, `@unpublished`  
✅ Online Resources: `@online`, `@electronic`, `@www` (with URL and access dates)  
✅ Other: `@patent`, `@booklet`, `@collection`, `@misc`

### Field Support
✅ **Standard fields**: `author`, `title`, `year`, `journal`, `publisher`, `pages`, etc.  
✅ **BibLaTeX fields**: `date` (auto-extracts year), `urldate`, `institution`, `type`  
✅ **URLs**: Automatically formatted with `[Accessed: date]`  
✅ **Special handling**: Curly brackets removed from all fields except URLs  
✅ **Author parsing**: Handles `{Acronyms}`, applies "et al." for >3 authors  
✅ **Page formatting**: Converts LaTeX double-dash (`123--145`) to single dash (`123-145`)

### Example

```bibtex
@article{smith2020,
  author = {John Smith and Jane Doe},
  title = {A Great Paper},
  journal = {Journal of Science},
  year = {2020},
  pages = {123--145}
}

@thesis{jones2023,
  author = {Jane Jones},
  title = {An Important Study},
  institution = {University of Example},
  type = {Ph.D. dissertation},
  date = {2023-05-15}
}

@misc{report2025,
  author = {Ministry of Research},
  title = {Annual Report},
  url = {https://example.com/report.pdf},
  urldate = {2026-02-07},
  date = {2025}
}
```

**Output formatting**:
- `[1] John Smith, Jane Doe. "A Great Paper". Journal of Science, pp. 123-145, 2020.`
- `[2] Jane Jones. "An Important Study". Ph.D. dissertation, University of Example, 2023.`
- `[3] Ministry of Research. "Annual Report". 2025. https://example.com/report.pdf [Accessed: 2026-02-07].`

## Example Workflow

1. **Create your presentation** with citation commands:
   - Slide 1: `"Recent research shows \cite{smith2020} that..."`
   - Slide 2: `"Multiple studies \cite{jones2019,doe2021} have found..."`
   - Last Slide: Title "References", content `\bibliography`

2. **Create a .bib file** (e.g., `presentation.bib`) with your sources

3. **Run the script**:
   ```bash
   python pptx_cite.py presentation.pptx
   ```

4. **Open `presentation_cited.pptx`** to see:
   - `\cite{smith2020}` → `[1]`
   - `\cite{jones2019,doe2021}` → `[2, 3]`
   - `\bibliography` → Full formatted reference list

## What the Tool Does

✅ Finds all `\cite{key}` commands in your presentation  
✅ Replaces them with formatted citations ([1], [2], etc.)  
✅ Generates complete bibliography at `\bibliography` location  
✅ Preserves all text formatting (bold, italic, colors, bullets)  
✅ Handles overflow with automatic continuation slides  
✅ Numbers citations by order of appearance, not alphabetically

## Requirements

- Python 3.7+
- python-pptx
- PyYAML

## Advanced Features

### Format Preservation
✅ Text formatting (bold, italic, font size, color)  
✅ Paragraph properties (bullets, numbering, indentation)  
✅ Run-level formatting within citations  
✅ All formatting maintained during citation replacement

### Bibliography Overflow Handling
✅ Automatically detects when content exceeds slide space  
✅ Creates continuation slides with identical formatting  
✅ Maintains consistent styling across all bibliography slides

### Smart Author & Field Processing
✅ Removes protective curly brackets: `{Smith}` → `Smith`  
✅ Uses "et al." for publications with >3 authors  
✅ Supports complex author name formats  
✅ URLs preserved with curly brackets intact  
✅ Converts page dashes: `123--145` → `123-145`

### Citation Order
✅ Citations numbered by order of first appearance  
✅ Not alphabetical - respects your presentation flow  
✅ Repeated citations maintain original number

## Testing

The `tests/` directory contains test files

## Limitations

⚠️ Text replacement preserves most formatting; complex nested styles may vary  
⚠️ Works with: text boxes, shapes, grouped shapes, and tables  
⚠️ **SmartArt**: Cannot be automatically processed due to PowerPoint's internal structure  
⚠️ Citations in charts may have limited support  
⚠️ Requires proper LaTeX-style citation syntax

### SmartArt Workaround

If you have citations in SmartArt graphics:
1. Right-click the SmartArt graphic
2. Select "Convert to Shapes" or "Group" → "Ungroup"
3. Run the script - it will now process the converted shapes
4. Alternatively, add citations in text boxes positioned near the SmartArt

## Command-Line Options

```
usage: pptx_cite.py [-h] [-c CONFIG] [-o OUTPUT] pptx_file [bib_file]

positional arguments:
  pptx_file             Path to the PowerPoint file (.pptx)
  bib_file              Path to the BibTeX file (.bib) [optional, defaults to pptx_file with .bib extension]

optional arguments:
  -h, --help            show this help message and exit
  -c CONFIG, --config CONFIG
                        Path to configuration file (optional)
  -o OUTPUT, --output OUTPUT
                        Output file path (optional, defaults to pptx_file with _cited suffix)
  -s, --start-number START_NUMBER
                        Starting number for numeric citation styles (optional, default: 1)
```

## License

MIT License - Feel free to use and modify as needed.
