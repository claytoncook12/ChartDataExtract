import argparse
import glob
import json
import sys
from datetime import datetime
from pathlib import Path

from extractor import load_excel, get_sheet_names, extract_cells


def save_output(content: str, output_format: str) -> str:
    """Save content to a file with timestamp. Returns the filename."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"extract_{timestamp}.{output_format}"

    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)

    return filename


def output_result(content: str, output_format: str):
    """Output result - save to file if format specified, otherwise print."""
    if output_format in ("csv", "json"):
        filename = save_output(content, output_format)
        print(f"Extraction complete. Data saved to: {filename}")
    else:
        print(content)


def format_output(results: dict, output_format: str = "text") -> str:
    """Format extraction results for display."""
    if output_format == "json":
        return json.dumps(results, indent=2, default=str)

    if output_format == "csv":
        lines = []
        # Check if results are nested (from config with sheet names)
        first_value = next(iter(results.values()), None)
        if isinstance(first_value, dict):
            # Nested: {sheet: {cell: value}}
            for sheet_name, cells in results.items():
                cell_refs = list(cells.keys())
                lines.append(",".join(["Sheet"] + cell_refs))
                row = [sheet_name] + [str(v) if v is not None else "" for v in cells.values()]
                lines.append(",".join(row))
        else:
            # Flat: {cell: value}
            cell_refs = list(results.keys())
            lines.append(",".join(cell_refs))
            row = [str(v) if v is not None else "" for v in results.values()]
            lines.append(",".join(row))
        return "\n".join(lines)

    # Text format
    lines = []
    for cell_ref, value in results.items():
        if isinstance(value, list):
            # Range result - format as table
            lines.append(f"{cell_ref}:")
            for row in value:
                lines.append("  " + " | ".join(str(v) if v is not None else "" for v in row))
        else:
            lines.append(f"{cell_ref}: {value}")
    return "\n".join(lines)


def format_table(batch_results: dict, cell_refs: list[str], output_format: str = "text") -> str:
    """Format batch results as a table with files as rows and cells as columns."""
    if output_format == "json":
        return json.dumps(batch_results, indent=2, default=str)

    if output_format == "csv":
        lines = []
        # Header
        lines.append(",".join(["File"] + cell_refs))
        # Data rows
        for filename, values in batch_results.items():
            row = [filename]
            for cell in cell_refs:
                val = values.get(cell, "")
                row.append(str(val) if val is not None else "")
            lines.append(",".join(row))
        return "\n".join(lines)

    # Text table format
    # Calculate column widths
    col_widths = {"File": max(len("File"), max(len(f) for f in batch_results.keys()))}
    for cell in cell_refs:
        max_val_len = 0
        for values in batch_results.values():
            val = values.get(cell, "")
            max_val_len = max(max_val_len, len(str(val) if val is not None else ""))
        col_widths[cell] = max(len(cell), max_val_len)

    # Build table
    lines = []
    # Header
    header = "| " + " | ".join(
        col.ljust(col_widths.get(col, len(col)))
        for col in ["File"] + cell_refs
    ) + " |"
    lines.append(header)

    # Separator
    sep = "|-" + "-|-".join("-" * col_widths.get(col, len(col)) for col in ["File"] + cell_refs) + "-|"
    lines.append(sep)

    # Data rows
    for filename, values in batch_results.items():
        row_data = [filename.ljust(col_widths["File"])]
        for cell in cell_refs:
            val = values.get(cell, "")
            val_str = str(val) if val is not None else ""
            row_data.append(val_str.ljust(col_widths[cell]))
        lines.append("| " + " | ".join(row_data) + " |")

    return "\n".join(lines)


def process_batch(file_patterns: list[str], sheet_name: str, cell_refs: list[str]) -> dict:
    """Process multiple files and return results keyed by filename."""
    results = {}

    for pattern in file_patterns:
        # Expand glob pattern
        files = glob.glob(pattern)
        if not files:
            print(f"Warning: No files match pattern '{pattern}'", file=sys.stderr)
            continue

        for file_path in files:
            path = Path(file_path)
            if not path.suffix.lower() == '.xlsx':
                continue

            try:
                workbook = load_excel(str(path))
                if sheet_name not in workbook.sheetnames:
                    print(f"Warning: Sheet '{sheet_name}' not found in {path.name}", file=sys.stderr)
                    continue

                sheet = workbook[sheet_name]
                file_results = extract_cells(sheet, cell_refs)
                results[path.name] = file_results
            except Exception as e:
                print(f"Warning: Error processing {path.name}: {e}", file=sys.stderr)

    return results


def main():
    parser = argparse.ArgumentParser(
        description="Extract values from Excel files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Single file
  python main.py data.xlsx --sheet Sheet1 --cells A1,B2,C3

  # Multiple files (table output)
  python main.py "test/*.xlsx" --sheet "Chart Data" --cells D1,D12,D120

  # Multiple patterns
  python main.py "folder1/*.xlsx" "folder2/*.xlsx" --sheet Sheet1 --cells A1,B2

  # CSV output
  python main.py "test/*.xlsx" --sheet "Chart Data" --cells D1,D12 --format csv

  # Using config
  python main.py "test/*.xlsx" --config config.json
        """
    )

    parser.add_argument("files", nargs="+", help="Path(s) to Excel file(s) - supports glob patterns like 'folder/*.xlsx'")
    parser.add_argument("--sheet", "-s", help="Sheet name to extract from")
    parser.add_argument("--cells", "-c", help="Comma-separated cell references (e.g., A1,B2,C3)")
    parser.add_argument("--list-sheets", action="store_true", help="List all sheet names and exit")
    parser.add_argument("--config", help="Path to JSON config file with extraction settings")
    parser.add_argument("--format", "-f", choices=["text", "json", "csv"], default="text",
                        help="Output format (default: text)")

    args = parser.parse_args()

    # Check if this is a batch operation (multiple files or glob pattern)
    expanded_files = []
    for pattern in args.files:
        expanded_files.extend(glob.glob(pattern))

    is_batch = len(expanded_files) > 1

    # Single file mode
    if not is_batch and len(expanded_files) == 1:
        file_path = Path(expanded_files[0])
        if not file_path.exists():
            print(f"Error: File not found: {expanded_files[0]}", file=sys.stderr)
            sys.exit(1)

        try:
            workbook = load_excel(str(file_path))
        except Exception as e:
            print(f"Error loading Excel file: {e}", file=sys.stderr)
            sys.exit(1)

        # List sheets mode
        if args.list_sheets:
            print("Sheets in workbook:")
            for name in get_sheet_names(workbook):
                print(f"  - {name}")
            return

        # Config file mode (single file)
        if args.config:
            config_path = Path(args.config)
            if not config_path.exists():
                print(f"Error: Config file not found: {args.config}", file=sys.stderr)
                sys.exit(1)

            with open(config_path) as f:
                config = json.load(f)

            all_results = {}
            for extraction in config.get("extractions", []):
                sheet_name = extraction.get("sheet")
                cells = extraction.get("cells", [])

                if sheet_name not in workbook.sheetnames:
                    print(f"Warning: Sheet '{sheet_name}' not found, skipping", file=sys.stderr)
                    continue

                sheet = workbook[sheet_name]
                results = extract_cells(sheet, cells)
                all_results[sheet_name] = results

            output_result(format_output(all_results, args.format), args.format)
            return

        # Direct extraction mode (single file)
        if not args.sheet:
            print("Error: --sheet is required (or use --list-sheets to see available sheets)", file=sys.stderr)
            sys.exit(1)

        if args.sheet not in workbook.sheetnames:
            print(f"Error: Sheet '{args.sheet}' not found", file=sys.stderr)
            print(f"Available sheets: {', '.join(get_sheet_names(workbook))}", file=sys.stderr)
            sys.exit(1)

        if not args.cells:
            print("Error: --cells is required", file=sys.stderr)
            sys.exit(1)

        sheet = workbook[args.sheet]
        cell_refs = [c.strip() for c in args.cells.split(",")]
        results = extract_cells(sheet, cell_refs)

        output_result(format_output(results, args.format), args.format)
        return

    # Batch mode - multiple files
    if not expanded_files:
        print(f"Error: No files found matching pattern(s)", file=sys.stderr)
        sys.exit(1)

    # Load config or use command line args
    if args.config:
        config_path = Path(args.config)
        if not config_path.exists():
            print(f"Error: Config file not found: {args.config}", file=sys.stderr)
            sys.exit(1)

        with open(config_path) as f:
            config = json.load(f)

        # Use first extraction from config for batch
        extraction = config.get("extractions", [{}])[0]
        sheet_name = extraction.get("sheet")
        cell_refs = extraction.get("cells", [])
    else:
        if not args.sheet:
            print("Error: --sheet is required for batch processing", file=sys.stderr)
            sys.exit(1)
        if not args.cells:
            print("Error: --cells is required for batch processing", file=sys.stderr)
            sys.exit(1)
        sheet_name = args.sheet
        cell_refs = [c.strip() for c in args.cells.split(",")]

    batch_results = process_batch(args.files, sheet_name, cell_refs)

    if not batch_results:
        print("No data extracted from any files", file=sys.stderr)
        sys.exit(1)

    output_result(format_table(batch_results, cell_refs, args.format), args.format)


if __name__ == "__main__":
    main()
