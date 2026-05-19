# Gebiet Card Generator

A Python tool that generates printable PNG cards from Excel data.

The script reads address data, “do not visit” entries, and optional extra notes from an Excel workbook and creates one PNG card per `Gebiet`.

## Features

- Reads data from an Excel file with multiple sheets
- Generates one PNG card per area / `Gebiet`
- Supports merged Excel cells
- Handles apartment ranges such as `1-100`
- Calculates total number of apartments per area
- Adds a separate “Bitte NICHT besuchen” section
- Supports optional extra information from an `Extras` sheet
- Automatically adjusts long house numbers to fit into the table
- Saves all generated cards as PNG files

## Project Structure

```text
.
├── generate_cards.py      # Main script
├── requirements.txt       # Python dependencies
├── .gitignore             # Files ignored by Git
└── README.md              # Project documentation
```

## Excel File Structure

The script expects an Excel workbook with the following sheets:

### `Daten`

Main address data.

Expected columns:

| Column | Description |
|---|---|
| `Gebiet` | Area number |
| `Straßen Name` | Street name |
| `Haus` | House number |
| `Stiegen` | Staircase / building section |
| `Parteien` | Apartments / units |
| `Parteien_calc` | Optional calculated apartment count |

### `NichtBesuchen`

Addresses that should not be visited.

Expected columns may include:

| Column | Description |
|---|---|
| `Gebiet` | Area number |
| `Straßen Name` | Street name |
| `Haus` | House number |
| `Stiegen` | Staircase |
| `Parteien` | Apartment / unit |
| `Top` | Apartment number |
| `Text` | Optional ready-made text line |

### `Extras`

Optional additional notes.

Expected columns:

| Column | Description |
|---|---|
| `Gebiet` | Area number |
| `Text` | Additional note |

## Installation

1. Clone this repository:

```bash
git clone https://github.com/YOUR_USERNAME/gebiet-card-generator.git
cd gebiet-card-generator
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Open `generate_cards.py` and change the Excel file path:

```python
EXCEL_FILE = "C:\\Users\\VivoBook\\Desktop\\Versammlung\\Data.xlsx"
```

Then run:

```bash
python generate_cards.py
```

The generated PNG files will be saved in the `output` folder.

## Output

The script creates files like:

```text
output/Gebiet_1.png
output/Gebiet_2.png
output/Gebiet_3.png
```

Each card contains:

- the Gebiet number,
- a table with street, house, staircase and apartment information,
- calculated total number of apartments,
- optional extra notes,
- a “Bitte NICHT besuchen” section.

## Technologies Used

- Python
- pandas
- openpyxl
- Pillow

## Possible Future Improvements

- Move settings to a separate config file
- Add command-line arguments
- Add a graphical user interface
- Export cards as PDF
- Add sample Excel data
- Add automated tests
