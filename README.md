# slide2pdf

A simple Python command-line tool that converts all PowerPoint `.ppt` and `.pptx` files in a folder into PDF format using Microsoft PowerPoint.

---

## Features

- Converts all `.ppt` and `.pptx` files in a folder to PDF
- Outputs PDFs into a subfolder named `p2pdf`
- Automatically uses current directory if no path is specified
- Works using Microsoft PowerPoint (must be installed)
- Windows-only (uses COM automation)

---

## Installation

> Requires Python 3.6+ and Microsoft PowerPoint installed on Windows.

### 1. Install from PyPI:

```bash
pip install slide2pdf
```

### 2. Install from source:

```bash
git clone https://github.com/yourusername/slide2pdf.git
cd slide2pdf
pip install .
```

---

## Usage

### Basic command (current directory):

```bash
slide2pdf
```

Finds and converts all PowerPoint files in the folder where you run the command.

### With a custom folder path:

```bash
slide2pdf --path "C:\Users\YourName\Documents\Slides"
```

Converts all .ppt and .pptx files in the specified folder.

### Output

All converted PDFs are saved inside a folder named `p2pdf` inside the input folder:

```
C:\Users\YourName\Documents\Slides\p2pdf\your_file.pdf
```

### Example

```bash
slide2pdf
# or
slide2pdf --path "D:\College\Sem6\Presentations"
```

---

## Requirements

- Microsoft PowerPoint installed (required for conversion)
- Windows OS
- Python 3.6+

## Important Note

This tool is designed to work exclusively on Windows systems. It requires Windows Subsystem for Linux (WSL) to function properly.

## Prerequisites

- Windows 10 or later
- Windows Subsystem for Linux (WSL) installed
- Python 3.x
- Microsoft PowerPoint installed

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For support, please [open an issue](https://github.com/AMV0027/slide2pdf) in the GitHub repository.
