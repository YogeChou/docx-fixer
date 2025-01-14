# docx-fixer
fix broken .docx file automatically

## Introduction

# Fix Word XML Errors

This repository provides a Python script for fixing corrupted or invalid XML files embedded in Microsoft Word `.docx` documents. The script automates the process of identifying, repairing, and replacing problematic XML content within Word files.

---

## Script Purpose

The main goal of this script is to recover and fix XML errors in `word/document.xml` files within Word documents. These errors often cause Word files to fail to open or render correctly. The script employs two approaches for error handling:

1. **Removing invalid characters**: Cleans the XML content by stripping out non-printable or invalid characters.
2. **Auto-fixing malformed XML**: Uses the `lxml` library to recover and repair XML structure automatically.
