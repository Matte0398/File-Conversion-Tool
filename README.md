# File-Conversion-Tool

Python script that converts CSV files to Excel and Excel files to CSV.

The script supports direct command-line arguments and configuration loading from a JSON file.

## Features

- Convert CSV files to Excel `.xlsx`
- Convert Excel `.xls` and `.xlsx` files to CSV
- Export one CSV file per Excel worksheet
- Process a specific list of files
- Process all supported files in a folder
- Custom CSV delimiter
- Optional overwrite mode
- JSON configuration file support
- Automatic output folders
- Operation log generation

## Requirements

- Python 3
- pandas
- openpyxl

## Installation

``` bash
pip install pandas openpyxl
```

## JSON configuration example

``` bash
{
  "path": "C:/data",
  "file_list": "ALL",
  "csv_delimiter": ";",
  "force_overwrite": true
}
```

## Usage

Convert all files inside the `/data` directory:

``` bash
python ExcelCSVConverter.py -P "/data" -L ALL
```

Convert only the specified files inside the `/data` directory:

``` bash
python ExcelCSVConverter.py -P "/data" -L "file1.csv,file2.xlsx"
```

Run the conversion using the parameters defined in `config.json`:

``` bash
python ExcelCSVConverter.py --config config.json
```

Convert all files inside the `/data` directory and force the conversion/overwrite behavior:

``` bash
python ExcelCSVConverter.py -P "/data" -L ALL -F
```

## Output

The script creates:

- `<path>`/CSV/ <br>
- `<path>`/XLSX/ <br>
- `<path>`/logConverter_<timestamp>.log
