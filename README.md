# Excel to PDF Converter

A React web application that converts Excel pallet tag files to PDF format, matching the exact layout and formatting of the sample output.

## Features

- Upload Excel files (.xlsx or .xls)
- Automatically parse Excel data
- Generate PDF with exact format matching the sample
- Download PDF files

## Installation

1. Install dependencies:
```bash
npm install
```

2. Start the development server:
```bash
npm run dev
```

3. Open your browser and navigate to the URL shown (usually `http://localhost:5173`)

## Usage

1. Click "Choose Excel File" and select your Excel file
2. Click "Convert to PDF"
3. The PDF will automatically download with the same name as your Excel file (with .pdf extension)

## Excel File Format

The app expects Excel files with the following columns (case-insensitive, flexible naming):
- LC / PO # (or LC/PO#, LC/PO, PO#)
- Customer (or Customer Name)
- Address (or Address Line 1)
- Country
- PROFORMA INVOICE NUMBER (or Invoice Number, Invoice)
- FILM DESCP (or Film Description, Description)
- SIZE MM (or Size, Size MM)
- No. OF REELS / PALLET (or Reels, Number of Reels)
- NET WEIGHT (PALLET) (or Net Weight, Net Weight (Pallet))
- GROSS WEIGHT (PALLET) (or Gross Weight, Gross Weight (Pallet))
- PALLET DIMENSIONS MM (or Dimensions)
- Length, Width, Height (individual columns, or combined in Dimensions)
- PALLET NUMBER (or Pallet Number, Pallet #)
- CONT # (or CONT#, Container Number) - can also be extracted from filename

## Technologies Used

- React 18
- Vite
- xlsx - for Excel file parsing
- jsPDF - for PDF generation

## Build for Production

```bash
npm run build
```

The built files will be in the `dist` folder.

