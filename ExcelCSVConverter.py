##################################################################
## Description: Script to convert an Excel to CSV or viceversa
##
## Author: Matteo Z.
##################################################################

import sys, argparse, json
import pandas as pd
from pathlib import Path
from datetime import datetime

def print_usage():
    usage = f"""
Usage:
    python {sys.argv[0]} [OPTIONS]

Description:
    Converts CSV files to Excel (.xlsx) and Excel files (.xls/.xlsx) to CSV (one CSV per worksheet)
    Supports both direct CLI arguments and configuration loading from a JSON file

Options:
    -P, --path <folder>
        Target folder containing the files to process
        Required unless provided via '--config'

    -L, --file-list <files>
        Comma-separated list of files (e.g. "a.csv,b.xlsx") or the  keyword 'ALL' to process every .csv/.xls/.xlsx file in the target folder
        Required unless provided via '--config'

    -D, --csv-delimiter <char>
        Optional - delimiter for CSV output (default: ",")
        May be defined in the JSON config file

    -F, --force-overwrite
        Optional - overwrite existing output files instead of skipping them

    --config <config.json>
        Optional - path to the JSON configuration file containing default values:

            {{
                "path": "C:/data",
                "file_list": "ALL",
                "csv_delimiter": ";",
                "force_overwrite": true
            }}

        CLI arguments always override values from the config file

    -h, --help
        Show this help message and exit

Examples:
    python {sys.argv[0]} -P "C:/data" -L ALL
    python {sys.argv[0]} -P "/data" -L "file1.csv,file2.xlsx"
    python {sys.argv[0]} --config config.json
    python {sys.argv[0]} --config config.json -D ";"
    python {sys.argv[0]} -P "/data" -L ALL -F

Output:
    - CSV files are saved in: <path>/CSV/
    - Excel files are saved in: <path>/XLSX/
    - Log file is saved in: <path>/logConverter_<timestamp>.log

"""
    print(usage)
    sys.exit(1)


class FileConverter:
    """Handles conversion between CSV and Excel formats"""
    
    def __init__(self, base_path: Path, csv_delimiter: str = ",", force_overwrite: bool = False):
        self.base_path = base_path
        self.csv_delimiter = csv_delimiter
        self.force_overwrite = force_overwrite
        self.log_entries = []
        
        # Create output folders
        self.csv_folder = base_path / "CSV"
        self.xlsx_folder = base_path / "XLSX"
        self.csv_folder.mkdir(parents=True, exist_ok=True)
        self.xlsx_folder.mkdir(parents=True, exist_ok=True)
    
    def get_files_to_process(self, file_list: str) -> list[Path]:
        """Returns the list of files to process"""
        extensions = {".csv", ".xls", ".xlsx"}
        
        if file_list.upper() == "ALL":
            return [f for f in self.base_path.iterdir() 
                   if f.is_file() and f.suffix.lower() in extensions]
        
        files = []
        for filename in file_list.split(","):
            filename = filename.strip()
            if not filename:
                continue
                
            filepath = self.base_path / filename
            if filepath.is_file():
                files.append(filepath)
            else:
                print(f"⚠️  File not found: {filename}", file=sys.stderr)
        
        return files
    
    def csv_to_excel(self, csv_file: Path) -> bool:
        """Converts a CSV file to Excel"""
        output_file = self.xlsx_folder / f"{csv_file.stem}.xlsx"
        
        if output_file.exists() and not self.force_overwrite:
            print(f"⏭️  Skipping (already exists): {output_file.name}")
            return False
        
        try:
            df = pd.read_csv(csv_file)
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Sheet1", index=False)
            
            msg = f"CSV → Excel: '{csv_file.name}' → '{output_file.name}'"
            self.log_entries.append(msg)
            print(f"✓ {msg}")
            return True
            
        except Exception as e:
            msg = f"Error converting '{csv_file.name}': {e}"
            self.log_entries.append(msg)
            print(f"✗ {msg}", file=sys.stderr)
            return False
    
    def excel_to_csv(self, excel_file: Path) -> int:
        """Converts an Excel file to CSV (one per sheet). Returns number of sheets converted"""
        converted = 0
        
        try:
            excel_data = pd.ExcelFile(excel_file)
        except Exception as e:
            msg = f"Error opening '{excel_file.name}': {e}"
            self.log_entries.append(msg)
            print(f"✗ {msg}", file=sys.stderr)
            return 0
        
        for sheet_name in excel_data.sheet_names:
            output_file = self.csv_folder / f"{excel_file.stem}_{sheet_name}.csv"
            
            if output_file.exists() and not self.force_overwrite:
                print(f"⏭️  Skipping (already exists): {output_file.name}")
                continue
            
            try:
                df = pd.read_excel(excel_data, sheet_name=sheet_name)
                
                if df.empty:
                    print(f"⏭️  Skipping empty sheet: '{sheet_name}' in '{excel_file.name}'")
                    continue
                
                df.to_csv(output_file, index=False, sep=self.csv_delimiter, encoding="utf-8")
                
                msg = f"Excel → CSV: '{excel_file.name}' [sheet: {sheet_name}] → '{output_file.name}'"
                self.log_entries.append(msg)
                print(f"✓ {msg}")
                converted += 1
                
            except Exception as e:
                msg = f"Error processing sheet '{sheet_name}' in '{excel_file.name}': {e}"
                self.log_entries.append(msg)
                print(f"✗ {msg}", file=sys.stderr)
        
        return converted
    
    def process_file(self, file: Path):
        """Processes a single file based on its extension"""
        ext = file.suffix.lower()
        
        if ext == ".csv":
            self.csv_to_excel(file)
        elif ext in {".xls", ".xlsx"}:
            self.excel_to_csv(file)
        else:
            msg = f"Unsupported file type: '{file.name}'"
            self.log_entries.append(msg)
            print(f"⚠️  {msg}")
    
    def save_log(self):
        """Saves the operations log"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self.base_path / f"logConverter_{timestamp}.log"
        
        lines = [
            "=" * 60,
            f"Conversion Log - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "=" * 60,
            ""
        ]
        
        for entry in self.log_entries:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            lines.append(f"[{ts}] {entry}")
        
        lines.extend([
            "",
            "=" * 60,
            f"Total operations: {len(self.log_entries)}",
            "=" * 60
        ])
        
        log_file.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n📄 Log saved to: {log_file}")


def load_config(config_path):
    """ to load the configuration from a JSON file """
    try:
        with open(config_path, encoding="utf-8") as f:
            config = json.load(f)

            if not isinstance(config, dict):
                print("⚠️ JSON file must contain an object at root level", file=sys.stderr)
                return {}
            
            return config
    except FileNotFoundError:
        print(f"⚠️ Configuration file not found: {config_path}", file=sys.stderr)
    except json.JSONDecodeError as e:
        print(f"⚠️ JSON parsing error: {e}", file=sys.stderr)
    except Exception as e:
        print(f"⚠️ Error reading configuration: {e}", file=sys.stderr)
    
    return {}


def main():
    parser = argparse.ArgumentParser(
        description="Convert CSV files to Excel or viceversa",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False
    )
    
    parser.add_argument("-P", "--path", help="Folder containing the files to process")
    parser.add_argument("-L", "--file-list", help="Comma-separated file list or 'ALL'")
    parser.add_argument("-D", "--csv-delimiter", help="CSV delimiter (default: ',')")
    parser.add_argument("-F", "--force-overwrite", action="store_true", help="Overwrite existing files")
    parser.add_argument("--config", help="JSON configuration file path")
    parser.add_argument("-h", "--help", action="store_true", help="Show detailed help message")
    args = parser.parse_args()
    
    if args.help:
        print_usage()

    config = load_config(args.config) if args.config else {}
    
    # Priority: CLI arguments > config file > defaults
    path = args.path or config.get("path")
    file_list = args.file_list or config.get("file_list")
    csv_delimiter = args.csv_delimiter or config.get("csv_delimiter", ",")
    force_overwrite = args.force_overwrite or config.get("force_overwrite", False)
    
    if not path or not file_list:
        print("❌ Error: Must specify --path and --file-list (or use --config)", file=sys.stderr)
        print(f"   Run 'python {sys.argv[0]} --help' for usage information\n", file=sys.stderr)
        return 1
    
    base_path = Path(path).expanduser().resolve()
    if not base_path.is_dir():
        print(f"❌ Error: '{base_path}' is not a valid directory", file=sys.stderr)
        return 1
    
    # Execute conversion
    print(f"\n🔄 Starting conversion in: {base_path}")
    print(f"   CSV delimiter: '{csv_delimiter}'")
    print(f"   Overwrite existing: {'Yes' if force_overwrite else 'No'}\n")
    
    converter = FileConverter(base_path, csv_delimiter, force_overwrite)
    files = converter.get_files_to_process(file_list)
    
    if not files:
        print("⚠️  No files to process")
        return 0
    
    print(f"📋 Files to process: {len(files)}\n")
    
    for file in files:
        converter.process_file(file)
    
    converter.save_log()
    print("\n✅ Conversion completed!")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
