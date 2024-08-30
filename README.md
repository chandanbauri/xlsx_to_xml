# `xlsx_to_xml`

`xlsx_to_xml` is a Python library that enables you to convert Excel files (.xlsx) into XML format easily and efficiently. This library is built with Rust, offering high performance and reliability, making it perfect for developers and data professionals who need to transform spreadsheet data into structured XML.

## Features

- **Effortless Conversion**: Convert entire Excel workbooks or specific sheets to XML.
- **High Performance**: Utilizes Rust for fast and efficient processing.
- **Customizable Output**: Supports customization to match various XML schemas.
- **Cross-Platform**: Compatible with Windows, macOS, and Linux.

## Installation

To install the library, use `pip`:

```bash
pip install xlsx-to-xml

```

Here’s an example of how to use xlsx_to_xml:

```python

from xlsx_to_xml import xlsx_to_xml

try:
    xml_output = xlsx_to_xml.convert_xlsx_to_xml("example.xlsx", "Sheet1")
    print(xml_output)
except Exception as e:
    print(f"An error occurred: {e}")

```


