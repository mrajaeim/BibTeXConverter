# BibTeX to Word XML Converter

A Python application that converts BibTeX files to Word Bibliography XML format, making it easy to import references into Microsoft Word documents.

## Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Quick Start (Windows Executable)](#quick-start-windows-executable)
- [Supported Reference Types](#supported-reference-types)
- [License](#license)

## Features

- User-friendly GUI interface
- Converts BibTeX (.bib) files to Word-compatible XML format
- Supports common reference types (articles, books, conference proceedings, etc.)
- Handles author names and basic bibliography fields

## Requirements

- Python 3.x
- bibtexparser >= 1.4.0


## Installation

1. Clone this repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python app.py
```

2. Click "Select BibTeX File" to choose your .bib file
3. Click "Convert to XML" to generate the Word-compatible XML
4. Choose where to save the output XML file
5. Import the generated XML file into Microsoft Word's bibliography manager

## Quick Start (Windows Executable)

For Windows users who want to use the application without installing Python:

1. Download the latest `BibTeXConverter.exe` from the releases section
2. Double-click the executable to run
3. No installation or dependencies needed

**⚠️ Disclaimer:** The executable is provided as-is, without any warranty. Use at your own risk. While we've created this executable for convenience, always exercise caution when running executable files from the internet.

## Supported Reference Types

- Journal Articles
- Books
- Conference Proceedings
- Theses/Reports

## License

This project is open source and available under the MIT License.
