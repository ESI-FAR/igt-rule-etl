# IGT Rules ETL

Scripts to convert between Excel used by human annotators in the HyChain WP3 institutional network analysis coding scheme, and JSON-like formats used by the [INA-tool](https://github.com/ESI-FAR/INA-tool).

## Overview

This tool provides functionality to convert rule and connection data between:
- Excel spreadsheet format (with "rules" and "connections" sheets - see examples in the ```testdata/``` folder)
- JSON-based format (JSON/JavaScript object notation - used by the INA-tool - see examples in the ```testdata/``` folder)

## Installation

1. Clone this repository
2. Ensure you have Python 3.6+ installed
3. Install the required dependencies:
   ```
   pip install pandas openpyxl
   ```
   or 

   ```
   pip install -r requirements.txt
   ```

## Command Line Usage

The tool provides a command-line interface for easy use:

### Converting Excel to JSON Format

```bash
python cli.py excel_to_json <excel_file> [--rules_output RULES_OUTPUT_FILENAME] [--connections_output CONNECTIONS_OUTPUT_FILENAME]
```

Arguments:

- `excel_file`: Path to the Excel file containing "rules" and "connections" sheets
- `--rules_output`: Path to save the rules text output (default: "rules_output.txt")
- `--connections_output`: Path to save the connections text output (default: "connections_output.txt")

Example:

```bash
python cli.py excel_to_json hychain_wp3_connections_ground_truth.xlsx --rules_output my_rules.txt --connections_output my_connections.txt
```

### Converting JSON to Excel Format

```bash
python cli.py json_to_excel <rules_file> <connections_file> [--excel_output EXCEL_OUTPUT_FILENAME]
```

Arguments:

- `rules_file`: Path to the rules text file
- `connections_file`: Path to the connections text file
- `--excel_output`: Path to save the Excel output (default: "converted_output.xlsx")

Example:

```bash
python cli.py json_to_excel example_rules.txt example_connections.txt --excel_output my_converted_data.xlsx
```

## Programmatic Usage

You can also use the conversion functions directly in your Python code:

```python
from core_conversion import excel_to_json_format, json_to_excel_format

# Convert Excel to JSON format
rules_json, connections_json = excel_to_json_format('hychain_wp3_connections_ground_truth.xlsx')

# Save to JSON string files
with open('rules_output.txt', 'w') as f:
    f.write(rules_json)

with open('connections_output.txt', 'w') as f:
    f.write(connections_json)

# Convert JSON files to Excel
json_to_excel_format('example_rules.txt', 'example_connections.txt', 'converted_output.xlsx')
```

## Input Format Details

### Excel Format

The Excel file should have two sheets:
- `rules`: Contains rule data with columns like "id", "Statement Type", "Attribute", etc. (see ```testdata/converted_output.xlsx``` for a full list of columns required)
- `connections`: Contains connection data with columns like "source_statement", "target_statement", etc. (see ```testdata/converted_output.xlsx``` for a full list of columns required)

### JSON Format

The input JSON strings should be placed in plain text files using a JSON/JavaScript object notation format:

#### Rules Format
```json
[
  {
    "Id": "1",
    "Statement": "...",
    "Statement Type": "formal",
    "Attribute": "insurers",
    "Deontic": "must",
    "Aim": "pay",
    "Direct Object": "property owner",
    "Type of Direct Object": "animate",
    "Indirect Object": "",
    "Type of Indirect Object": "",
    "Activation Condition": "if named storm and damage",
    "Execution Constraint": "",
    "Or Else": ""
  },
  ...
]
```

#### Connections Format
```json
[
  {
    "driven_by": "actor",
    "source_component": "Direct Object",
    "source_statement": "6",
    "target_component": "Attribute",
    "target_statement": "12"
  },
  ...
]
```

## License

The contents of this repository is licensed under the MIT License - see the [LICENSE file](https://raw.githubusercontent.com/ESI-FAR/igt-rule-etl/refs/heads/main/LICENSE) for details.