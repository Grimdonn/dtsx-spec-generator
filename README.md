# DTSX Spec Document Generator

A Python-based tool to automatically generate Microsoft Word specification documents from SQL Server Integration Services (SSIS) `.dtsx` package files.

## Overview

Documenting SSIS packages manually can be time-consuming and prone to errors. This script parses the XML structure of a `.dtsx` file to extract critical information and formats it into a professional Word document, serving as a technical specification or "Spec Doc".

## Key Features

- **Connection Managers**: Extracts name, connection type, and connection strings (e.g., OLE DB, Flat File).
- **Variables**: Identifies both System and User variables, including their data types and default values.
- **SQL Code Extraction**: Automatically pulls SQL queries from OLE DB Sources and Lookup transformation components.
- **Flat File Schema**: Documents flat file column definitions, including data types, widths, and delimiters.
- **Professional Output**: Generates a structured `.docx` file with clear headings and tables.

## Prerequisites

- **Python 3.x**
- **python-docx**: Library for creating and updating Microsoft Word files.

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/Grimdonn/dtsx-spec-generator.git
   cd dtsx-spec-generator
   ```

2. Install the required dependencies:
   ```bash
   pip install python-docx
   ```

## Usage

1. Place your `.dtsx` file in the project directory.
2. Open `dtsx_parser.py` and set the `dtsx_path` and `output_docx` variables:
   ```python
   dtsx_path = "YourPackage.dtsx"
   output_docx = "Package_Spec_Doc.docx"
   ```
3. Run the script:
   ```bash
   python dtsx_parser.py
   ```
4. The generated specification document will be available as `Package_Spec_Doc.docx`.

## Example Output Sections

The generated document includes the following sections:
- **Package Information**: High-level metadata about the source file.
- **Connection Managers**: A tabular view of all connections.
- **Variables**: Categorized list of variables and their attributes.
- **Extracted SQL Code**: Dedicated sections for each component containing SQL logic.
- **Flat File Column Definitions**: Detailed schemas for every flat file connection.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request or open an issue for any bugs or feature requests.

