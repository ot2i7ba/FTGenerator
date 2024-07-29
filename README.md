# Fake Timesheet Generator
Ever found yourself forgetting to fill out timesheets for months on end? You're not alone! That's why I, a self-confessed lazy developer, created this nifty little script. The `timesheet.py` generates random fake work schedules and timesheets for a given month, including Excel and PDF exports, based on user-defined parameters.

> [!WARNING]
> This script produces entirely fake data based on the variables you set. Depending on how these documents are used, the recipient may consider them false, invalid, or even fraudulent. Use it wisely and with caution!

For my purposes, this script is a lifesaver, helping me keep up with paperwork I otherwise can't be bothered to do. It’s tailored to suit my needs perfectly, and it is completely legitimate for my own use. If you're mature enough to understand the implications, feel free to adapt it for your own purposes. Just remember: with great laziness comes great responsibility!

## How It Works
When you run the script, you'll be prompted to enter several parameters: the year and month for which you want to generate the timesheet, the minimum and maximum number of workdays, the earliest start time and latest end time for work daily hours, and the minimum and maximum total hours for the month. Based on these inputs, the script randomly selects workdays within the specified range, ensuring they do not fall on holidays or weekends. It then generates random start and end times for each workday within the provided time window, ensuring that the total hours worked fall within the specified limits. The generated data is then compiled into detailed timesheets, which are saved in CSV and Excel formats, with an option to include a digital signature and export the document as a PDF.

## Features
- **Random Work Schedule Generation**<br>Because who needs real schedules? Generates random work schedules for any given month, ensuring your fake work times look varied and oh-so realistic.
- **CSV, Excel, and PDF Exports**<br>Creates detailed timesheets in CSV and Excel formats, with a fancy PDF export for easy sharing and pretending. Impress your friends and colleagues with beautifully organized fake data.
- **Customizable Parameters**<br>Lets you play god with your work hours: define the number of workdays, start and end times, and total working hours. Flexibility is key when you're making stuff up!
- **German Holidays Consideration**<br>Because even fake employees need their holidays! Automatically skips German public holidays and special days (like Christmas Eve and New Year's Eve) to keep your schedule "realistic."
- **Signature Support**<br>Add a personal touch with your very own signature and a snazzy image of it in the Excel timesheets. The image is scaled to fit perfectly, maintaining your artistic integrity.
- **Error Handling and Retry Mechanism**<br>No more infinite loops! Ensures valid timesheet generation with multiple retries in case your made-up hours don’t quite add up. It's like having a built-in sanity check for your fakery.
- **User-Friendly Prompts**<br>Interactive prompts guide even the laziest user through the setup process. Customizing your fake timesheets has never been easier!
- **Directory Management**<br>Automatically creates and organizes output files in neatly named directories, because even our laziness has standards.
- **Existing File Check**<br>Prevents accidental overwriting of your precious fake data. Gives you the option to back up or overwrite existing files – because nobody wants to lose their masterpiece.
- **Bulk Processing Option**<br>Got a whole year to fake? No problem! Supports bulk processing to generate timesheets for all months in one go. More fakery, less effort.
- **Data Preview**<br>Preview your brilliantly concocted timesheet data in the terminal before committing to it. Verify and tweak your fake schedules to perfection.

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
```sh
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
The generated timesheets will be saved in a directory named `Stundenzettel {year}`, with files in CSV, Excel, and PDF formats.

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

# Screenshot
<img src="https://github.com/ot2i7ba/FTGenerator/blob/main/assets/images/screenshot.png" width="32%" alt="Fake Timesheet Generator"> 

___

# License
This project is licensed under the **[MIT license](https://github.com/ot2i7ba/FTGenerator/blob/main/LICENSE)**, providing users with flexibility and freedom to use and modify the software according to their needs.

# Contributing
Contributions are welcome! Please fork the repository and submit a pull request for review.

# Disclaimer
This project is provided without warranties. Users are advised to review the accompanying license for more information on the terms of use and limitations of liability.

# Conclusion
Let's be honest: I'm too lazy to fill out timesheets regularly. In fact, I've been known to forget them for months, sometimes even years! Instead of manually creating Excel tables and saving them as PDFs, I built this handy little helper. It's tailored to my needs, which is why the generated files are in German. But don't worry, you can customize the script to suit your own preferences. So, if you're as forgetful (or lazy) as I am, this script might just save your day!

