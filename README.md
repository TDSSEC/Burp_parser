# Burp XML Parser

This Python script parses Burp Suite XML files containing security issues and exports the data to either CSV or XLSX format. It extracts specific fields from the XML file, such as issue details, and organizes them into a tabular format suitable for further analysis.

## Why?
I struggled to find a working burpsuite parser that worked with the latest burpsuite pro 2024 edition, so here it is.

## Prerequisites

- Python 3.x
- Required Python packages:
  - `xml.etree.ElementTree`
  - `openpyxl`

Install the necessary packages using pip:

```bash
pip install openpyxl
```

## Usage  
1. Clone the Repository
```bash
git clone https://github.com/TDSSEC/burp-xml-parser.git
```
2. Navigate to the Repository
```
cd burp-xml-parser
```
3. Run the Script
```
python3 burp_parser.py <path_to_burp_xml_file> <output_file.csv|output_file.xlsx>
```
Replace <path_to_burp_xml_file> with the path to your Burp Suite XML file and <output_file.csv|output_file.xlsx> with the desired output filename (ending with .csv or .xlsx).
```
python3 burp_parser.py burp_export.xml output.csv
```
4. Check Output  
After running the script, check the generated output file (output.csv or output.xlsx) for the extracted data.

## Additional Notes
The script automatically detects the output file format based on the filename extension (.csv or .xlsx).  
If the output file is an XLSX file, a formatted Excel table (ListObject) is created with a predefined style (TableStyleMedium2).
