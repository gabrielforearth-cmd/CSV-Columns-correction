# CSV and Excel Column Formatter

`fix_columns.py` is a Python utility that repairs CSV or Excel files whose content was incorrectly loaded into a single column.

The script detects the actual delimiter, separates the values into their correct columns, cleans unnecessary quotation marks, and exports the result as a formatted Excel workbook.

## Features

* Supports CSV, XLS, and XLSX files
* Automatically detects common delimiters
* Repairs files imported as a single column
* Respects quoted CSV values
* Removes unnecessary quotation marks
* Trims leading and trailing whitespace
* Preserves empty values
* Exports the result as an XLSX file
* Displays the detected columns and total row count

## Supported Delimiters

The script can detect the following delimiters:

* Comma: `,`
* Semicolon: `;`
* Tab: `\t`
* Pipe: `|`

Delimiter detection first uses `csv.Sniffer`.

When automatic detection fails, the script counts the occurrences of each supported delimiter and selects the most frequent one.

## Requirements

* Python 3.8 or later
* pandas
* openpyxl

Install the required packages with:

```bash
pip install pandas openpyxl
```

For legacy `.xls` files, an additional Excel engine may be required:

```bash
pip install xlrd
```

## Project Structure

```text
column-formatter/
├── fix_columns.py
└── README.md
```

## Usage

Run the script from the command line and provide the path to the input file:

```bash
python fix_columns.py path/to/file.csv
```

For an Excel file:

```bash
python fix_columns.py path/to/file.xlsx
```

Example on Windows:

```bash
python fix_columns.py "C:\Data\Customers.csv"
```

Example on Linux:

```bash
python fix_columns.py "/home/user/data/Customers.csv"
```

## Output

The generated file is saved in the same directory as the input file.

The output filename follows this format:

```text
<original_filename>_formatted.xlsx
```

Examples:

```text
Customers.csv
Customers_formatted.xlsx
```

```text
Partners.xlsx
Partners_formatted.xlsx
```

The original file is not modified.

## Complete Script

```python
#!/usr/bin/env python3
"""
fix_columns.py

Usage:
    python fix_columns.py path/to/file.csv
    python fix_columns.py path/to/file.xlsx

Result:
    Generates: path/to/file_formatted.xlsx
"""

import csv
import os
import sys

import pandas as pd


def detect_delimiter(sample):
    """
    Detects the delimiter using csv.Sniffer.

    If automatic detection fails, the function selects the delimiter
    with the highest number of occurrences.
    """
    supported_delimiters = [",", ";", "\t", "|"]

    try:
        dialect = csv.Sniffer().sniff(
            sample,
            delimiters=supported_delimiters
        )
        return dialect.delimiter

    except Exception:
        counts = {
            delimiter: sample.count(delimiter)
            for delimiter in supported_delimiters
        }

        return max(counts, key=counts.get)


def parse_csv_line(line, delimiter):
    """
    Parses one CSV line while respecting quoted values.
    """
    return next(
        csv.reader(
            [line],
            delimiter=delimiter,
            quotechar='"'
        )
    )


def fix_single_column_df(dataframe, delimiter=None):
    """
    Repairs a DataFrame in which every cell in the first column
    contains an entire CSV line.

    The first parsed row is used as the column header.
    """
    first_column = dataframe.iloc[:, 0].astype(str).tolist()
    sample = "\n".join(first_column[:10])

    if delimiter is None:
        delimiter = detect_delimiter(sample)

    parsed_rows = [
        parse_csv_line(line, delimiter)
        for line in first_column
    ]

    parsed_dataframe = pd.DataFrame(parsed_rows)

    parsed_dataframe.columns = (
        parsed_dataframe.iloc[0].astype(str)
    )

    parsed_dataframe = (
        parsed_dataframe[1:]
        .reset_index(drop=True)
    )

    parsed_dataframe.columns = [
        column.strip().replace('"', "")
        for column in parsed_dataframe.columns
    ]

    parsed_dataframe = parsed_dataframe.map(
        lambda value: (
            str(value).strip().replace('"', "")
            if pd.notnull(value)
            else value
        )
    )

    return parsed_dataframe, delimiter


def main(infile, outfile=None, encoding="utf-8"):
    if not os.path.isfile(infile):
        print("File not found:", infile)
        return

    base, extension = os.path.splitext(infile)

    if outfile is None:
        outfile = f"{base}_formatted.xlsx"

    extension = extension.lower()

    if extension in (".xls", ".xlsx"):
        dataframe = pd.read_excel(
            infile,
            header=None,
            dtype=str
        )

    else:
        with open(
            infile,
            "r",
            encoding=encoding,
            errors="replace"
        ) as file:
            sample_lines = "".join(
                file.readline()
                for _ in range(10)
            )

        delimiter = detect_delimiter(sample_lines)

        try:
            dataframe = pd.read_csv(
                infile,
                sep=delimiter,
                quotechar='"',
                encoding=encoding,
                dtype=str,
                engine="python"
            )

        except Exception:
            with open(
                infile,
                "r",
                encoding=encoding,
                errors="replace"
            ) as file:
                lines = [
                    line.rstrip("\n\r")
                    for line in file
                ]

            dataframe = pd.DataFrame(lines)

    if dataframe.shape[1] == 1:
        fixed_dataframe, used_delimiter = (
            fix_single_column_df(dataframe)
        )

        fixed_dataframe.to_excel(
            outfile,
            index=False
        )

        print("File saved at:", outfile)
        print(
            "Detected delimiter:",
            repr(used_delimiter)
        )
        print(
            "Columns:",
            list(fixed_dataframe.columns)
        )
        print(
            "Rows:",
            fixed_dataframe.shape[0]
        )

        return

    dataframe.columns = [
        str(column).strip().replace('"', "")
        for column in dataframe.columns
    ]

    dataframe = dataframe.map(
        lambda value: (
            str(value).strip().replace('"', "")
            if pd.notnull(value)
            else value
        )
    )

    dataframe.to_excel(
        outfile,
        index=False
    )

    print("File saved at:", outfile)
    print("Columns:", list(dataframe.columns))
    print("Rows:", dataframe.shape[0])


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(
            "Usage: python fix_columns.py "
            "path/to/file.csv or path/to/file.xlsx"
        )
    else:
        main(sys.argv[1])
```

## How It Works

### 1. Input validation

The script checks whether the provided input path points to an existing file:

```python
if not os.path.isfile(infile):
    print("File not found:", infile)
    return
```

If the file does not exist, processing stops.

### 2. File type identification

The input type is determined from its extension:

```python
base, extension = os.path.splitext(infile)
extension = extension.lower()
```

The following extensions are handled as Excel files:

```text
.xls
.xlsx
```

All other extensions are processed as delimited text files.

### 3. CSV delimiter detection

For CSV files, the script reads the first ten lines:

```python
with open(
    infile,
    "r",
    encoding=encoding,
    errors="replace"
) as file:
    sample_lines = "".join(
        file.readline()
        for _ in range(10)
    )
```

The sample is passed to `detect_delimiter`.

The function first tries:

```python
csv.Sniffer().sniff(...)
```

If that fails, it counts each supported delimiter:

```python
counts = {
    delimiter: sample.count(delimiter)
    for delimiter in supported_delimiters
}
```

The delimiter with the highest occurrence count is selected.

### 4. Normal CSV loading

The script initially attempts to load the CSV directly with pandas:

```python
dataframe = pd.read_csv(
    infile,
    sep=delimiter,
    quotechar='"',
    encoding=encoding,
    dtype=str,
    engine="python"
)
```

Using `dtype=str` prevents pandas from automatically converting identifiers into numbers.

For example:

```text
00001234
```

remains:

```text
00001234
```

instead of becoming:

```text
1234
```

### 5. Raw-line fallback

If pandas cannot load the CSV normally, the script reads each physical line as a single value:

```python
lines = [
    line.rstrip("\n\r")
    for line in file
]
```

These lines are then stored in a one-column DataFrame.

### 6. Single-column detection

The script checks the number of columns:

```python
if dataframe.shape[1] == 1:
```

A one-column result may indicate that each cell contains an entire CSV row.

Example of incorrect structure:

| Column 1                      |
| ----------------------------- |
| `"Number";"Name";"Country"`   |
| `"100";"Company A";"France"`  |
| `"200";"Company B";"Germany"` |

The script parses each row and converts it into:

| Number | Name      | Country |
| ------ | --------- | ------- |
| 100    | Company A | France  |
| 200    | Company B | Germany |

### 7. Quoted-value parsing

Each line is processed with Python's CSV parser:

```python
csv.reader(
    [line],
    delimiter=delimiter,
    quotechar='"'
)
```

This is important because quoted fields may contain delimiters.

Example:

```csv
"100","Company, Incorporated","United States"
```

The comma inside `Company, Incorporated` is treated as part of the field rather than as a column separator.

### 8. Header reconstruction

When repairing a single-column file, the first parsed row becomes the header:

```python
parsed_dataframe.columns = (
    parsed_dataframe.iloc[0].astype(str)
)
```

The first row is then removed from the dataset:

```python
parsed_dataframe = (
    parsed_dataframe[1:]
    .reset_index(drop=True)
)
```

### 9. Data cleaning

Column names and values are cleaned by:

* Removing leading whitespace
* Removing trailing whitespace
* Removing remaining quotation marks

Example:

```text
" Customer number "
```

becomes:

```text
Customer number
```

Data values receive the same cleanup.

### 10. Excel export

The final DataFrame is exported using:

```python
dataframe.to_excel(
    outfile,
    index=False
)
```

The pandas index is not added to the generated workbook.

## Example

### Input file

```csv
"Number";"Customer Name";"Country"
"0001";"Company A";"France"
"0002";"Company B";"Germany"
```

When incorrectly loaded, the file may appear as a single column:

```text
"Number";"Customer Name";"Country"
"0001";"Company A";"France"
"0002";"Company B";"Germany"
```

### Generated workbook

| Number | Customer Name | Country |
| ------ | ------------- | ------- |
| 0001   | Company A     | France  |
| 0002   | Company B     | Germany |

### Terminal output

```text
File saved at: Customers_formatted.xlsx
Detected delimiter: ';'
Columns: ['Number', 'Customer Name', 'Country']
Rows: 2
```

## Already Formatted Files

When the input file already contains multiple columns, the script does not reconstruct the structure.

It only:

1. Cleans the column names
2. Cleans the cell values
3. Removes unnecessary quotation marks
4. Exports the result to XLSX

Example output:

```text
File saved at: Partners_formatted.xlsx
Columns: ['Number', 'Name', 'Customer number']
Rows: 350
```

## Character Encoding

The default CSV encoding is:

```text
UTF-8
```

It is defined by:

```python
def main(infile, outfile=None, encoding="utf-8"):
```

Invalid byte sequences are replaced because the file is opened using:

```python
errors="replace"
```

This prevents the script from stopping because of isolated invalid characters. However, replaced characters may appear as:

```text
�
```

The script does not currently perform automatic encoding detection.

To process another encoding programmatically:

```python
main(
    "Customers.csv",
    encoding="cp1252"
)
```

For a Big5-encoded file:

```python
main(
    "Customers.csv",
    encoding="big5"
)
```

## Custom Output Path

The command-line interface automatically creates the output path.

When importing the script as a module, a custom path can be provided:

```python
from fix_columns import main

main(
    infile="Customers.csv",
    outfile="results/Customers_cleaned.xlsx"
)
```

## Important Notes

### Original file preservation

The source file is never overwritten unless the same path is explicitly provided as `outfile` when calling `main` programmatically.

### Identifier preservation

Using `dtype=str` helps preserve:

* Leading zeros
* Customer numbers
* Partner numbers
* Invoice identifiers
* Purchase-order identifiers
* Codes containing numeric characters

### Quotation-mark removal

The cleanup removes all double quotation marks from values:

```python
str(value).strip().replace('"', "")
```

This means legitimate quotation marks inside textual content will also be removed.

For example:

```text
Product "Premium"
```

becomes:

```text
Product Premium
```

### Excel header behavior

Excel files are loaded with:

```python
header=None
```

This prevents pandas from automatically discarding or interpreting the first row as a header.

When the Excel workbook contains a single column, the first row is promoted to the header during reconstruction.

When the workbook already contains multiple columns, the script retains pandas-generated numeric column names such as:

```text
0
1
2
3
```

The first worksheet row remains part of the data.

### First worksheet only

For Excel files, pandas reads the first worksheet by default.

The script does not process every worksheet in the workbook.

### Uneven row lengths

If reconstructed CSV lines contain different numbers of values, pandas fills missing positions with empty values.

Example:

```csv
"100","Company A","France"
"200","Company B"
```

Result:

| Column 1 | Column 2  | Column 3 |
| -------- | --------- | -------- |
| 100      | Company A | France   |
| 200      | Company B |          |

## Error Handling

The script directly handles a missing input file:

```text
File not found: path/to/file.csv
```

When no path is provided, it displays:

```text
Usage: python fix_columns.py path/to/file.csv or path/to/file.xlsx
```

The CSV reader also includes a fallback when normal pandas parsing fails.

Errors related to the following conditions may still be raised:

* Invalid Excel workbooks
* Password-protected Excel files
* Missing Excel engine dependencies
* Permission errors
* Invalid output directories
* Files currently locked by another application
* Rows with incompatible CSV structures

## Limitations

* Automatic encoding detection is not implemented.
* The default CSV encoding is UTF-8.
* Invalid characters may be replaced silently.
* Only the first Excel worksheet is processed.
* Legitimate double quotation marks are removed.
* Delimiter detection is based on a sample of the first ten lines.
* Excel files with multiple columns do not automatically promote the first row to headers.
* The command-line interface accepts only the input path.
* Custom output paths and encodings require calling `main` programmatically.
* The output is always an XLSX workbook.
* Formula values and Excel formatting may not be preserved.

## Recommended Workflow

1. Keep a backup of the original file.
2. Close the file in Excel before running the script.
3. Run `fix_columns.py` with the source path.
4. Review the detected delimiter.
5. Review the displayed column names.
6. Confirm the reported number of rows.
7. Open the generated XLSX workbook.
8. Verify that identifiers and special characters were preserved.
9. Use the formatted workbook for the next processing or import step.

## License

This utility is intended for internal data formatting, spreadsheet correction, and CSV normalization operations.
