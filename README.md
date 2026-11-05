# ChartDataExtract

Extract cell values from Excel files via command line.

## Setup

```bash
# Activate virtual environment
.\venv\Scripts\activate
```

## Usage

```bash
# Single file
python main.py data.xlsx --sheet "Sheet1" --cells "A1,B2,C3"

# Multiple files (table output)
python main.py "folder/*.xlsx" --sheet "Sheet1" --cells "A1,B2,C3"

# Using config file
python main.py "folder/*.xlsx" --config config.json

# List sheets in a file
python main.py data.xlsx --list-sheets
```

## Output Formats

```bash
# Terminal output (default)
python main.py "test/*.xlsx" --config config.json

# Save to CSV file
python main.py "test/*.xlsx" --config config.json --format csv

# Save to JSON file
python main.py "test/*.xlsx" --config config.json --format json
```

CSV/JSON outputs are saved as `extract_YYYYMMDD_HHMMSS.{csv|json}`

## Config File

```json
{
  "extractions": [
    {
      "sheet": "Chart Data",
      "cells": ["D1", "D12", "D120", "C1", "C12", "C120"]
    }
  ]
}
```

## Options

| Option | Description |
|--------|-------------|
| `--sheet`, `-s` | Sheet name to extract from |
| `--cells`, `-c` | Comma-separated cell references (e.g., A1,B2,C3:C10) |
| `--config` | Path to JSON config file |
| `--format`, `-f` | Output format: text, csv, json |
| `--list-sheets` | List all sheet names |
