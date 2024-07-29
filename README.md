# Fake Timesheet Generator
This script, generates random fake work schedules and timesheets for a given month, including Excel and PDF exports, based on user-defined parameters.

## Features
- **Random Work Schedule Generation**<br>Generates random work schedules for any given month, ensuring realistic and varied work times.
- **CSV, Excel, and PDF Exports**<br>Creates detailed timesheets in CSV and Excel formats, with PDF export for easy sharing and printing.
- **Customizable Parameters**<br>Allows users to define the number of workdays, start and end times, and total working hours.
- **German Holidays Consideration**<br>Automatically considers German public holidays and special days (e.g., Christmas Eve) to avoid generating work hours on these days.
- **Signature Support**<br>Includes an option to add a signature and an image of the signature in the generated Excel timesheets.
- **Error Handling and Retry Mechanism**<br>Ensures valid timesheet generation with multiple retries in case of conflicts with the defined constraints.

## Requirements
- Python 3.6 or higher
    - `openpyxl`
    - `holidays`
    - `pypiwin32` (for Windows PDF conversion)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-username/fake-timesheet-generator.git
cd fake-timesheet-generator
```

2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage
Run the script using Python:
```bash
python timesheet.py
```

## Configuration
Follow the on-screen prompts to configure your desired parameters:

- **Bulk Processing**<br>Option to generate timesheets for all months of a year.
- **Year**<br>The year for which the timesheet is generated.
- **Month**<br>The month for which the timesheet is generated (only if bulk processing is not selected).
- **Minimum and Maximum Workdays**<br>Define the range of working days in the month.
- **Start and End Times**<br>Specify the earliest start time and latest end time for work hours.
- **Minimum and Maximum Total Hours**<br>Set the range for the total working hours in the month.
- **Signature**<br>Enter a name for the signature and optionally choose an image file for the signature.
- **Preview**<br>Option to preview the generated schedule in the terminal.

# Example
```bash
python timesheet.py
```

## Sample Prompts:
```bash
Do you want to perform bulk processing? (yes/no) [no]:
Enter the year [2024]:
Enter the month (1-12) [7]:
Enter the maximum workdays in the month [8]:
Enter the minimum workdays in the month [6]:
Enter the earliest start time (HH:MM) [17:00]:
Enter the latest end time (HH:MM) [22:00]:
Enter the maximum total hours in the month [17]:
Enter the minimum total hours in the month [15]:
Enter the name for the signature:
The following PNG files were found:
1. signature1.png
2. signature2.png
Enter the number of the image to use for the signature (1-2), or press Enter to skip:
```

## Output
The generated timesheets will be saved in a directory named Stundenzettel {year}, with files in CSV, Excel, and PDF formats.

## File Structure
- **CSV File**<br>Contains the dates, start times, end times, and total hours for each workday.
- **Excel File**<br>Includes formatted timesheet with headers, total hours, and optional signature.
- **PDF File**<br>PDF version of the Excel timesheet for easy sharing and printing.

## Signature and Image
- **Signature Name**<br>The name entered for the signature will be added to the generated Excel timesheet.
- **Signature Image**<br>Optionally, you can add an image of your signature to the timesheet.
    - The image must be in PNG format.
    - It is recommended to use an image with good dimensions for better clarity. For example, a signature image of 215x80 pixels works well.
    - The script will attempt to scale the image to fit within the cell, keeping the aspect ratio intact. However, it is best to use an image with appropriate dimensions to avoid excessive scaling and maintain legibility.

___

# License
This project is licensed under the **[MIT license](https://github.com/ot2i7ba/FTGenerator/blob/main/LICENSE)**, providing users with flexibility and freedom to use and modify the software according to their needs.

# Contributing
Contributions are welcome! Please fork the repository and submit a pull request for review.

# Disclaimer
This project is provided without warranties. Users are advised to review the accompanying license for more information on the terms of use and limitations of liability.

# Conclusion
Let's be honest: I'm too lazy to fill out timesheets regularly. In fact, I've been known to forget them for months, sometimes even years! Instead of manually creating Excel tables and saving them as PDFs, I built this handy little helper. It's tailored to my needs, which is why the generated files are in German. But don't worry, you can customize the script to suit your own preferences. So, if you're as forgetful (or lazy) as I am, this script might just save your day!

