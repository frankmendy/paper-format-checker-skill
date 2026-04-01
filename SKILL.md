---
name: "paper_format_checker"
description: "Automatically fixes research paper formatting based on '网络空间安全学院论文格式模板.docx'. Invoke when user asks to fix, format, or align a paper with the template."
---

# Paper Format Fixer

## Goal
- Automatically adjust a research paper (.docx) to match the formatting of the official '网络空间安全学院论文格式模板.docx'.
- The fixed version will be saved as a new file with the suffix `_fixed`.

## When To Use
- The user asks to fix, format, or correct a paper's layout according to the template.
- The user wants the document margins, line spacing, and styles to be automatically updated for specific sections (Abstract, TOC, Body).

## Inputs
- Absolute path to the target document file (.docx).

## Steps
1. Ask the user for the absolute file path to the paper to fix.
2. Run the fixer script:
   - Windows:
     - `python D:\Trae\paper_format_checker\scripts\paper_checker.py --file "<ABSOLUTE_FILE_PATH>"`
3. The script will perform a three-part structured fix:
   - **Part 1: Abstract Section**: Fixes formatting for Title, Abstract, and Keywords.
   - **Part 2: TOC Section**: Fixes formatting for Table of Contents levels (toc 1, 2, 3).
   - **Part 3: Body & Captions**: Fixes formatting for Headings (1-3), body text, image alignments, and table/figure captions.
4. Report the completion and provide the path to the fixed file.

## Notes
- Original file remains untouched.
- Only `.docx` files are supported.
- Captions for figures must be below the image, and table captions must be above the table (script ensures standard alignment).

