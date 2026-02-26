# Forms Feedback Parser

## Description

This is a small Python script designed to parse Microsoft Feedback Forms for Butler Scrum courses.

## Installation

Prerequisites: Python >= v3.11

Run `pip install -r requirements.txt` inside the project directory to install all external libraries.

## Usage

Run the script using `python parser.py C:\path\to\form_spreadsheet.xlsx` while in the project directory.

The script will create one or many output spreadsheets for different customizable filters.

### Filters

Filter options are placed at the top of parser.py file.

The options include a date filter and a sprint number filter.

The only form with parsing implemented currently is the SE361 sprint review form.
