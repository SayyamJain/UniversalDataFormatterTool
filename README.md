# Data Formatter Tool

A Python-based GUI tool for converting data files between different formats, including JSON, CSV, XML, Excel, and YAML. This tool simplifies the process of reformatting data files for users with minimal technical background or those seeking a convenient solution to handle various data formats.


## Features

- **Easy File Conversion:** Converts data files between JSON, CSV, XML, Excel, and YAML formats.

- **Intuitive GUI:** A user-friendly interface built with Tkinter, making it accessible for non-technical users.

- **Real-time Status Updates:** Displays the current status of the conversion process.

- **Error Handling:** Built-in error handling for file reading and writing issues.


## Requirements

- **Python 3.6+**
- **Required Python packages:**

  - json (standard library)
  - csv (standard library)
  - xml.etree.ElementTree (standard library)
  - openpyxl (for Excel file operations)
  - yaml (for YAML file operations)
  - tkinter (for GUI)


## Installation

**1. Clone the repository:**
```
git clone https://github.com/yourusername/data-formatter-tool.git
cd data-formatter-tool
```

**2. Install the required packages:**
```
pip install openpyxl pyyaml
```

**3. Run the application:**
```
python universal_data_formatter.py
```

## Usage

1. Open the tool by running
   *data_formatter_tool.py*.

2. Select your Input File and Input Format.

3. Choose your desired Output File location and Output Format.

4. Click Convert to initiate the conversion.

5. The Status label will update in real-time, showing the progress of the conversion.


## File Formats Supported

- JSON
- CSV
- XML
- Excel (.xlsx)
- YAML


## Example

Convert data.json to data.csv:

1. In the GUI, select data.json as the Input File.

2. Set JSON as the Input Format.

3. Choose a destination for data.csv and select CSV as the Output Format.

4. Click Convert and wait for the status to show Conversion complete.


## Project Structure
```
data-formatter-tool/
├── data_formatter_tool.py   # Main application script
├── README.md                # Project documentation
```

## License

This project is open-source and available under the MIT License.


## Contributing

Feel free to submit issues, fork the repository, and send pull requests. All contributions are welcome!
