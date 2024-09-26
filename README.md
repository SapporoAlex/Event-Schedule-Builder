# Event Schedule Builder
This Python program allows teachers to quickly and neatly build event schedules, designed with simplicity and flexibility in mind. It enables users to input event details, add activities with their times, and export the final schedule as an Excel file. The program is available in both English and Japanese versions to accommodate different language preferences.

## Features
- Create New Event Schedules: Input key details such as event title, date, leader, location, equipment, and activities.
- Activity Management: Easily add multiple activities and times to the event schedule.
- Edit and Confirm: Review or edit all entered information before saving the final version.
- Excel Export: Automatically generates a clean and formatted Excel file from the provided event details, based on a customizable template.
- Wrap Text Formatting: Ensures that all content fits neatly into cells by enabling text wrapping for longer entries.

## How to Use
Run the Program: Launch the script and select "Build new schedule" to start creating an event schedule.
Enter Event Details: Input event information like title, date, leader, location, and equipment.
Add Activities: Input the times and details of activities to build the schedule.
Review & Edit: Check the entered information and make any necessary changes through the confirmation menu.
Save the File: Once everything is confirmed, save the schedule as an Excel file (.xlsx) with your preferred file name.

## Installation
Clone the repository:
```bash
git clone https://github.com/your-username/event-schedule-builder.git
```

Install dependencies:
```bash
pip install openpyxl
Run the Python script:
```

```bash
python "Event Builder.py"
```

## Excel Template
The program uses an Excel template (template.xlsx) stored in the template/ folder. You can customize this template to suit your specific needs, such as modifying cell locations or adding additional fields.

## Requirements
Python 3.x
openpyxl
Example Usage

```bash
$ python "Event Builder.py"
Follow the prompts to input event details and activities, and the program will generate an Excel file based on your inputs.
```

## Japanese Version
A Japanese version of the program is also included for convenience. The steps and features are the same, with the user interface fully translated into Japanese.

Enjoy using this tool to streamline your event scheduling process!

Feel free to adjust any details like the repository name or features to better fit your project!
