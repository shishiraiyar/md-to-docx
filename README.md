# md-to-docx

A simple PowerShell script to convert Markdown files to DOCX format with support for Mermaid diagrams.

## Requirements

This tool requires the following to be installed:

- **Mermaid CLI** (`mmdc`) - for converting Mermaid diagrams to images
- **Pandoc** - for converting Markdown to Word format

## Usage

```powershell
.\convert.ps1 -InputFile <path-to-markdown-file> -OutputFile <path-to-docx-file>
```

Or using aliases:

```powershell
.\convert.ps1 -i input.md -o output.docx
```
