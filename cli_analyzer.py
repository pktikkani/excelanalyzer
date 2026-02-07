"""
Command-line Excel Analyzer
Analyze Excel files with multiple sheets using pandas.
Usage: python cli_analyzer.py file1.xlsx file2.xlsx
"""

import os
import sys
from typing import Dict

import numpy as np
import pandas as pd


class CLIExcelAnalyzer:
    def __init__(self):
        self.dataframes: Dict[str, Dict[str, pd.DataFrame]] = {}
        self.file_info: Dict[str, Dict] = {}

    def load_excel_file(self, filepath: str) -> bool:
        """Load an Excel file and extract all sheets."""
        if not os.path.exists(filepath):
            print(f"  ✗ File not found: {filepath}")
            return False

        filename = os.path.basename(filepath)
        try:
            excel_file = pd.ExcelFile(filepath)
            sheets = {}

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(filepath, sheet_name=sheet_name)
                sheets[sheet_name] = df

            self.dataframes[filename] = sheets
            self.file_info[filename] = {
                "sheets": list(sheets.keys()),
                "total_sheets": len(sheets),
                "total_rows": sum(len(df) for df in sheets.values()),
                "total_columns": sum(len(df.columns) for df in sheets.values()),
            }

            print(f"  ✓ Loaded '{filename}' — {len(sheets)} sheet(s)")
            return True
        except Exception as e:
            print(f"  ✗ Error loading '{filepath}': {e}")
            return False

    def show_file_summary(self, filename: str):
        """Print summary information about a loaded file."""
        if filename not in self.file_info:
            print(f"File '{filename}' not found.")
            return

        info = self.file_info[filename]
        print(f"\n{'=' * 60}")
        print(f"  File: {filename}")
        print(f"  Sheets: {info['total_sheets']}")
        print(f"  Sheet Names: {', '.join(info['sheets'])}")
        print(f"  Total Rows: {info['total_rows']}")
        print(f"  Total Columns: {info['total_columns']}")
        print(f"{'=' * 60}")

    def show_sheet_info(self, filename: str, sheet_name: str):
        """Print detailed information about a specific sheet."""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            print("Sheet not found.")
            return

        df = self.dataframes[filename][sheet_name]
        print(f"\n  Sheet: {sheet_name} (from {filename})")
        print(f"  Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        print(f"  Columns: {', '.join(df.columns.tolist())}")
        print(f"  Data Types:")
        for col in df.columns:
            print(f"    - {col}: {df[col].dtype}")

        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            print(f"\n  Numeric Column Stats:")
            for col in numeric_cols:
                print(
                    f"    - {col}: min={df[col].min():.2f}, "
                    f"max={df[col].max():.2f}, "
                    f"mean={df[col].mean():.2f}"
                )

    def show_head(self, filename: str, sheet_name: str, n: int = 5):
        """Show first n rows of a sheet."""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            print("Sheet not found.")
            return

        df = self.dataframes[filename][sheet_name]
        print(f"\n  First {n} rows of '{sheet_name}' from '{filename}':\n")
        print(df.head(n).to_string(index=True))

    def show_tail(self, filename: str, sheet_name: str, n: int = 5):
        """Show last n rows of a sheet."""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            print("Sheet not found.")
            return

        df = self.dataframes[filename][sheet_name]
        print(f"\n  Last {n} rows of '{sheet_name}' from '{filename}':\n")
        print(df.tail(n).to_string(index=True))

    def show_describe(self, filename: str, sheet_name: str):
        """Show descriptive statistics for a sheet."""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            print("Sheet not found.")
            return

        df = self.dataframes[filename][sheet_name]
        print(f"\n  Descriptive Statistics for '{sheet_name}' from '{filename}':\n")
        print(df.describe(include="all").to_string())

    def show_nulls(self, filename: str, sheet_name: str):
        """Show null value counts for a sheet."""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            print("Sheet not found.")
            return

        df = self.dataframes[filename][sheet_name]
        null_counts = df.isnull().sum()
        print(f"\n  Null Value Counts for '{sheet_name}' from '{filename}':")
        for col in df.columns:
            count = null_counts[col]
            pct = (count / len(df)) * 100 if len(df) > 0 else 0
            marker = " ⚠" if count > 0 else ""
            print(f"    - {col}: {count} ({pct:.1f}%){marker}")

    def show_unique(self, filename: str, sheet_name: str, column: str):
        """Show unique values for a column."""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            print("Sheet not found.")
            return

        df = self.dataframes[filename][sheet_name]
        if column not in df.columns:
            print(f"  Column '{column}' not found in sheet '{sheet_name}'.")
            print(f"  Available columns: {', '.join(df.columns.tolist())}")
            return

        unique_vals = df[column].unique()
        print(f"\n  Unique values in '{column}' ({len(unique_vals)} total):")
        for val in unique_vals[:50]:
            print(f"    - {val}")
        if len(unique_vals) > 50:
            print(f"    ... and {len(unique_vals) - 50} more")

    def show_value_counts(self, filename: str, sheet_name: str, column: str):
        """Show value counts for a column."""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            print("Sheet not found.")
            return

        df = self.dataframes[filename][sheet_name]
        if column not in df.columns:
            print(f"  Column '{column}' not found in sheet '{sheet_name}'.")
            print(f"  Available columns: {', '.join(df.columns.tolist())}")
            return

        counts = df[column].value_counts()
        print(f"\n  Value counts for '{column}':")
        print(counts.head(30).to_string())
        if len(counts) > 30:
            print(f"  ... and {len(counts) - 30} more values")

    def compare_files(self):
        """Compare all loaded files side by side."""
        if len(self.dataframes) < 2:
            print("  Need at least 2 files to compare.")
            return

        print(f"\n{'=' * 60}")
        print("  File Comparison")
        print(f"{'=' * 60}")
        for filename, info in self.file_info.items():
            print(f"\n  {filename}:")
            print(f"    Sheets: {info['total_sheets']} — {', '.join(info['sheets'])}")
            print(f"    Total Rows: {info['total_rows']}")
            print(f"    Total Columns: {info['total_columns']}")

        # Find common sheet names
        all_sheet_names = [set(info["sheets"]) for info in self.file_info.values()]
        common_sheets = set.intersection(*all_sheet_names) if all_sheet_names else set()
        if common_sheets:
            print(f"\n  Common Sheet Names: {', '.join(common_sheets)}")
            for sheet in common_sheets:
                print(f"\n  Comparing sheet '{sheet}' across files:")
                for filename in self.dataframes:
                    df = self.dataframes[filename][sheet]
                    print(f"    {filename}: {df.shape[0]} rows × {df.shape[1]} columns")
        else:
            print("\n  No common sheet names found between files.")

    def filter_data(
        self, filename: str, sheet_name: str, column: str, operator: str, value: str
    ):
        """Filter data based on a condition."""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            print("Sheet not found.")
            return

        df = self.dataframes[filename][sheet_name]
        if column not in df.columns:
            print(f"  Column '{column}' not found.")
            print(f"  Available columns: {', '.join(df.columns.tolist())}")
            return

        try:
            # Try to convert value to number if possible
            try:
                numeric_value = float(value)
            except ValueError:
                numeric_value = None

            if operator == "==" or operator == "=":
                if numeric_value is not None and pd.api.types.is_numeric_dtype(
                    df[column]
                ):
                    mask = df[column] == numeric_value
                else:
                    mask = df[column].astype(str) == value
            elif operator == "!=" and numeric_value is not None:
                mask = df[column] != numeric_value
            elif operator == ">" and numeric_value is not None:
                mask = df[column] > numeric_value
            elif operator == ">=" and numeric_value is not None:
                mask = df[column] >= numeric_value
            elif operator == "<" and numeric_value is not None:
                mask = df[column] < numeric_value
            elif operator == "<=" and numeric_value is not None:
                mask = df[column] <= numeric_value
            elif operator == "contains":
                mask = df[column].astype(str).str.contains(value, case=False, na=False)
            else:
                print(
                    f"  Unsupported operator '{operator}' or non-numeric value for comparison."
                )
                return

            filtered = df[mask]
            print(f"\n  Filtered results ({len(filtered)} rows):\n")
            if len(filtered) > 0:
                print(filtered.head(50).to_string(index=True))
                if len(filtered) > 50:
                    print(f"\n  ... showing first 50 of {len(filtered)} rows")
            else:
                print("  No rows match the filter condition.")

        except Exception as e:
            print(f"  Error filtering: {e}")

    def _pick_file(self) -> str:
        """Interactive file picker."""
        filenames = list(self.dataframes.keys())
        if len(filenames) == 0:
            print("  No files loaded.")
            return None
        if len(filenames) == 1:
            return filenames[0]

        print("\n  Available files:")
        for i, name in enumerate(filenames, 1):
            print(f"    {i}. {name}")
        choice = input("  Select file number: ").strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(filenames):
                return filenames[idx]
        except ValueError:
            pass
        print("  Invalid selection.")
        return None

    def _pick_sheet(self, filename: str) -> str:
        """Interactive sheet picker."""
        sheets = list(self.dataframes[filename].keys())
        if len(sheets) == 1:
            return sheets[0]

        print(f"\n  Available sheets in '{filename}':")
        for i, name in enumerate(sheets, 1):
            print(f"    {i}. {name}")
        choice = input("  Select sheet number: ").strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(sheets):
                return sheets[idx]
        except ValueError:
            pass
        print("  Invalid selection.")
        return None

    def _pick_column(self, filename: str, sheet_name: str) -> str:
        """Interactive column picker."""
        df = self.dataframes[filename][sheet_name]
        columns = df.columns.tolist()
        print(f"\n  Available columns:")
        for i, col in enumerate(columns, 1):
            print(f"    {i}. {col} ({df[col].dtype})")
        choice = input("  Select column number: ").strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(columns):
                return columns[idx]
        except ValueError:
            pass
        print("  Invalid selection.")
        return None

    def print_help(self):
        """Print available commands."""
        print("""
╔══════════════════════════════════════════════════════════════╗
║                  Excel Analyzer — Commands                   ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║  load <filepath>      Load an Excel file                     ║
║  files                List all loaded files                   ║
║  summary              Show summary of a file                  ║
║  info                 Show detailed sheet info                ║
║  head                 Show first rows of a sheet              ║
║  tail                 Show last rows of a sheet               ║
║  describe             Show descriptive statistics             ║
║  nulls                Show null value counts                  ║
║  unique               Show unique values in a column          ║
║  counts               Show value counts for a column          ║
║  filter               Filter rows by condition                ║
║  compare              Compare loaded files                    ║
║  help                 Show this help message                  ║
║  quit / exit          Exit the analyzer                       ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
        """)

    def interactive_loop(self):
        """Run the interactive question-answer loop."""
        print("\n" + "=" * 60)
        print("  🔍 Excel File Analyzer (CLI)")
        print("  Type 'help' for available commands")
        print("=" * 60)

        while True:
            try:
                user_input = input("\n📊 > ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break

            if not user_input:
                continue

            parts = user_input.split(maxsplit=1)
            command = parts[0].lower()
            args = parts[1] if len(parts) > 1 else ""

            if command in ("quit", "exit", "q"):
                print("Goodbye!")
                break

            elif command == "help":
                self.print_help()

            elif command == "load":
                if not args:
                    args = input("  Enter file path: ").strip()
                if args:
                    self.load_excel_file(args)

            elif command == "files":
                if not self.dataframes:
                    print("  No files loaded. Use 'load <filepath>' to load a file.")
                else:
                    print("\n  Loaded Files:")
                    for fname in self.dataframes:
                        info = self.file_info[fname]
                        print(
                            f"    • {fname} — {info['total_sheets']} sheet(s), {info['total_rows']} rows"
                        )

            elif command == "summary":
                filename = self._pick_file()
                if filename:
                    self.show_file_summary(filename)

            elif command == "info":
                filename = self._pick_file()
                if filename:
                    sheet = self._pick_sheet(filename)
                    if sheet:
                        self.show_sheet_info(filename, sheet)

            elif command == "head":
                filename = self._pick_file()
                if filename:
                    sheet = self._pick_sheet(filename)
                    if sheet:
                        n_input = input("  Number of rows (default 5): ").strip()
                        n = int(n_input) if n_input.isdigit() else 5
                        self.show_head(filename, sheet, n)

            elif command == "tail":
                filename = self._pick_file()
                if filename:
                    sheet = self._pick_sheet(filename)
                    if sheet:
                        n_input = input("  Number of rows (default 5): ").strip()
                        n = int(n_input) if n_input.isdigit() else 5
                        self.show_tail(filename, sheet, n)

            elif command == "describe":
                filename = self._pick_file()
                if filename:
                    sheet = self._pick_sheet(filename)
                    if sheet:
                        self.show_describe(filename, sheet)

            elif command == "nulls":
                filename = self._pick_file()
                if filename:
                    sheet = self._pick_sheet(filename)
                    if sheet:
                        self.show_nulls(filename, sheet)

            elif command == "unique":
                filename = self._pick_file()
                if filename:
                    sheet = self._pick_sheet(filename)
                    if sheet:
                        column = self._pick_column(filename, sheet)
                        if column:
                            self.show_unique(filename, sheet, column)

            elif command == "counts":
                filename = self._pick_file()
                if filename:
                    sheet = self._pick_sheet(filename)
                    if sheet:
                        column = self._pick_column(filename, sheet)
                        if column:
                            self.show_value_counts(filename, sheet, column)

            elif command == "filter":
                filename = self._pick_file()
                if filename:
                    sheet = self._pick_sheet(filename)
                    if sheet:
                        column = self._pick_column(filename, sheet)
                        if column:
                            print("  Operators: ==, !=, >, >=, <, <=, contains")
                            operator = input("  Enter operator: ").strip()
                            value = input("  Enter value: ").strip()
                            self.filter_data(filename, sheet, column, operator, value)

            elif command == "compare":
                self.compare_files()

            else:
                print(
                    f"  Unknown command: '{command}'. Type 'help' for available commands."
                )


def main():
    analyzer = CLIExcelAnalyzer()

    # Load any files passed as arguments
    if len(sys.argv) > 1:
        print("Loading files...")
        for filepath in sys.argv[1:]:
            analyzer.load_excel_file(filepath)

    analyzer.interactive_loop()


if __name__ == "__main__":
    main()
