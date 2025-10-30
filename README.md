# Excel Splitter

A Python script that processes Excel files by splitting them into smaller CSV files and converting them to JSON format based on configurable environment variables.

## Features

- Loads Excel files from a specified source folder
- Splits large DataFrames into smaller CSV files with configurable row limits
- Converts entire DataFrame to JSON with column names as keys
- Creates organized output structure with dedicated folders for each Excel file
- Configurable through environment variables

## Requirements

- Python 3.6+
- pandas
- numpy
- openpyxl
- python-dotenv

## Installation

Install the required dependencies:

```bash
pip install pandas numpy openpyxl python-dotenv
```

## Configuration

The script uses environment variables defined in the `.env` file:

```env
SOURCE_FOLDER=./source_files
OUTPUT_FOLDER=./output
SPLIT_ROWS_MAX=200000
SPLIT_SUFFIX=part
```

### Environment Variables

- **SOURCE_FOLDER**: Directory containing Excel files to process (default: `./source_files`)
- **OUTPUT_FOLDER**: Directory where output files will be saved (default: `./output`)
- **SPLIT_ROWS_MAX**: Maximum number of rows per CSV split file (default: `200000`)
- **SPLIT_SUFFIX**: Suffix to use for split CSV filenames (default: `part`)

## Usage

### Option 1: Manual Execution

1. Place your Excel files in the source folder specified in `.env`
2. Run the script:

```bash
python excel_splitter.py
```

### Option 2: Automated Run Scripts

For a complete automated workflow that includes package installation, script execution, and output zipping, use the provided run scripts:

**Linux/macOS:**
```bash
./run.sh
```

**Windows:**
```cmd
run.bat
```

The run scripts will:
1. Install required Python packages
2. Execute `excel_splitter.py`
3. Zip each folder created in the output directory

## Output Structure

The script creates the following output structure:

```
output/
├── [excel_filename]/
│   ├── [excel_filename]_part1.csv
│   ├── [excel_filename]_part2.csv
│   ├── ...
│   └── [excel_filename].json
```

### Example

For an Excel file named `data.xlsx` with `SPLIT_ROWS_MAX=1000` and `SPLIT_SUFFIX=part`:

```
output/
├── data/
│   ├── data_part1.csv    # First 1000 rows
│   ├── data_part2.csv    # Next 1000 rows
│   ├── data_part3.csv    # Remaining rows
│   └── data.json         # Complete data as JSON
```

## Output Formats

### CSV Files
- Each CSV file contains a header row with column names
- Files are named using the pattern: `[original_filename]_[suffix][number].csv`
- Maximum rows per file determined by `SPLIT_ROWS_MAX`

### JSON File
- Column names become top-level keys
- Each column contains an array of all values from that column
- Named with the same base name as the original Excel file

## Script Behavior

1. **Environment Loading**: Loads configuration from `.env` file
2. **File Discovery**: Automatically finds the first Excel file in the source folder
3. **Data Loading**: Loads Excel file into a pandas DataFrame
4. **Output Folder Creation**: Creates a dedicated folder for each Excel file
5. **CSV Splitting**: Divides DataFrame into smaller CSV files based on row limit
6. **JSON Conversion**: Converts entire DataFrame to JSON format
7. **Progress Reporting**: Provides detailed progress information and summary

## Error Handling

The script includes error handling for:
- Missing Excel files in source folder
- Excel file reading errors
- File writing permissions
- Invalid environment variable values

## Example Output

```
Loading environment variables...
Source folder: ./source_files
Output folder: ./output
Split rows max: 10
Split suffix: part

Looking for Excel files in ./source_files...
Found Excel file: ./source_files/bulk2.xlsx
Loading Excel file into DataFrame...
Loaded DataFrame with shape: (29, 20)

Creating output folder...
Created output folder: ./output/bulk2

Splitting DataFrame into CSV files...
Splitting 29 rows into 3 files with max 10 rows each
Created: bulk2_part1.csv (rows 1-10)
Created: bulk2_part2.csv (rows 11-20)
Created: bulk2_part3.csv (rows 21-29)

Converting DataFrame to JSON...
Created: bulk2.json

=== Processing Complete ===
Original file: ./source_files/bulk2.xlsx
DataFrame shape: (29, 20)
Output folder: ./output/bulk2
CSV files created: 3
  - bulk2_part1.csv
  - bulk2_part2.csv
  - bulk2_part3.csv
JSON file created: bulk2.json
```

## File Structure

```
.
├── excel_splitter.py    # Main script
├── run.sh              # Linux/macOS run script
├── run.bat             # Windows run script
├── .env                # Environment variables
├── README.md           # This documentation
├── source_files/       # Input Excel files
│   └── data.xlsx
└── output/             # Generated output files
    ├── data/
    │   ├── data_part1.csv
    │   ├── data_part2.csv
    │   └── data.json
    └── data.zip        # Zipped output folder
```

## License

This project is open source and available under the MIT License.
