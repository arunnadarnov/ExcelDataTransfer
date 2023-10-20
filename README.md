# Excel Data Transfer

This project contains a Python script for transferring data from one Excel file to another.

## Description

The `ExcelDataTransfer` class in the script reads a JSON configuration file, which specifies the source and destination Excel files and the sheets within them. The script then reads data from the source file and writes it to the destination file.

## Directory Structure

The project should have the following structure:
```markdown
* Application/
    * ConfigFiles/
        * config.json
    * ExcelDataTransfer.py
```

## Getting Started

### Dependencies

* Python 3.x
* xlwings
* pandas
* json
* os

### Installing

1. Clone this repository.
2. Install the required Python packages using pip:

`pip install xlwings pandas`

### Configuration

Update the `config.json` file in the `ConfigFiles` directory with your source and destination files and sheet names. Here's an example of what the `config.json` file might look like:

[
    {
        "source_file": "Source File Location Here",
        "destination_file": "Destination file location here",
        "destination_sheet": "Data"
    }
]

In this example, data is read from the first sheet in `ShiftReport 0313202306-00.xlsx` and written to the `Data` sheet in `ACD Data Analysis Tools 2023A.xlsm` file.

### Executing program
1. Update the config.json file with your source and destination files and sheet names.
2. Run ExcelDataTransfer.py:
`python ExcelDataTransfer.py`

## Version History

* 0.1
    * Initial Release
