# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a KR converter application - a web-based tool for converting KIR (Knjiga izdanih računov / Issued invoices book) and KPR (Knjiga prejetih računov / Received invoices book) Excel files into FURS-compatible JSON and XML formats for Slovenian tax reporting. The application is a single-page HTML application with JavaScript that processes Excel files in the browser.

## Development Commands

- **Development server**: `bun run start` - starts hot-reloading development server
- **Production build**: `bun run prod` - installs dependencies and builds for production
- **Manual build**: `bun run build.js` - runs the build script directly

## Architecture

### Core Components

1. **Frontend (HTML/JS)**: Single-page application (`index.html` + `script.js`)
   - Form-based interface for metadata input (tax number, periods, checkboxes)
   - File upload handling for Excel files (.xls/.xlsx)
   - Real-time preview tables showing parsed data
   - Download links for generated JSON/XML/ZIP files

2. **Excel Processing**: Client-side parsing using XLSX.js library
   - `parse_kir()`: Parses issued invoices (starting from row 9, columns A-AC)
   - `parse_kpr()`: Parses received invoices (starting from row 10, columns A-Y)
   - Data parsing utilities for different types (integers, floats, dates, strings)

3. **Output Generation**:
   - `generate_furs_json()`: Creates FURS-compliant JSON structure
   - `generate_furs_xml()`: Converts JSON to XML with proper FURS schema
   - ZIP file creation using JSZip library

### Key Data Flow

1. User uploads Excel files (KIR/KPR) and fills metadata form
2. Files are parsed into structured objects with validation
3. Data is combined with metadata to generate export formats
4. User can preview data in tables and download files

### Excel File Structure

**KIR (Knjiga izdanih računov)**:
- Data starts at row 9
- Metadata in cells N1 (OBDOBJE) and S1 (OBRAVNAVA)
- Supports samoprijava fields (columns AD-AE)

**KPR (Knjiga prejetih računov)**:
- Data starts at row 10  
- Metadata in cells B1 (OBDOBJE) and D1 (OBRAVNAVA)
- Supports samoprijava fields (columns X-Y)

### Build System

The project uses Bun as the build tool with a custom build script (`build.js`) that:
- Processes HTML entry point
- Injects version timestamps into `script.js`
- Outputs to `_site` directory
- Uses a custom plugin to replace version placeholders

### Dependencies

- **XLSX**: Excel file parsing
- **JSZip**: ZIP file generation
- **date-fns**: Date formatting utilities
- **Bun**: Runtime and build tool