# Donation Report Generator

This Python project generates a detailed Excel report of monthly donation transactions. The program processes a CSV file containing transaction details, filters incoming payments, and produces an Excel file summarizing the data for each month in both German and Serbian Cyrillic. It also generates a summary sheet with total monthly incomes.

## Features

- **Language support**: Handles German and Serbian Cyrillic.
- **Currency formatting**: Converts localized currency formats (e.g., `1.234,56 €`) into numeric values.
- **Excel report generation**: Automatically generates an Excel file with detailed donation data, including names, amounts, purposes, and countries.
- **Automated initials and name respelling**: Converts specific letter patterns (e.g., `ic` to `ić`) and extracts initials for donor names.
- **Monthly summaries**: Compiles and formats summaries for each month and a total income report.

## Prerequisites

Before running the script, ensure you have the following installed:

- Python 3.x
- Pandas
- Openpyxl

To install the necessary dependencies, run:

```bash
pip install pandas openpyxl
````

## File Structure

- kontobericht{year}.csv: The input CSV file that contains the raw bank transaction data. The file should be named with the specific year, e.g., kontobericht2022.csv.

- Izvestaji/Donatori - Spender {year}.xlsx: The output Excel file, generated after processing the CSV data.

## Script Details

### Functions 

1. respell_serbian_name(name):
    - Corrects and replaces characters in Serbian names (like replacing 'ic' with 'ić' and 'dj' with 'đ').

2. extract_initials(name):
    - Extracts the initials from donor names, used to anonymize personal information in the report.

3. map_country(code):
    - Maps country codes (e.g., 'AT' for Austria, 'DE' for Germany) to full country names. If a code is missing, it defaults to 'Austria'.

4. convert_month_to_serbian_cyrillic(month_name):
    - Converts month names from German to Serbian Cyrillic.

5. delocalize(string):
    - Converts localized German number strings to floats, handling both commas and periods correctly.

## Workflow 

1. *Preparing and Preprocessing Data:*
    The script reads the input CSV file using pandas, renames columns to fit the required format, and filters the transactions where the amount (Betrag) is positive, i.e., only donations are included.

2. *Summing Donations by Month:*
    For each month, the script calculates the total donations and formats the data into a human-readable table. The sum for each month is written to a DataFrame called gesamt_daten, which later gets exported into the "GESAMT" sheet of the Excel file.

3. *Writing to Excel:*
    The script uses the *openpyxl* library to write data into multiple sheets, one for each month, in both Serbian and German. It adds titles, borders, formatted currency, and merges cells for a clean, professional look.

4. *Styling:*
    The script formats cells with specific fonts, fills, and alignments. It ensures that titles are bold and centered, borders are applied to all cells, and the currency values are aligned to the right with proper formatting.


## Output Excel File Structure
- *Monthly Sheets*
    Each month (e.g., Januar, Februar) has a corresponding sheet that contains the donation data for that month, formatted in both Serbian Cyrillic and German.

    *Headers* 
    - Име / Name (Name of the donor)
    - Земља / Land (Country)
    - Износ / Betrag (Amount)
    - Сврха / Zweck (Purpose)

    The sum of donations for each month is displayed at the bottom of each sheet.

- *GESAMT {year} Sheet:*
    Contains a summary of all donations received throughout the year, with the total amount clearly indicated.

## Example Usage    

To generate the report for the year 2022, make sure you have a file named kontobericht2022.csv in the correct format. Then, run the script:

```bash
python donatori_report.py   
```

## Error Handling 
- *AttributeError* or *IndexError*: These are caught and printed with the month name if any issues arise during the monthly processing loop.

- *Empty Data:* If a month has no data, it is skipped from the report without causing an error.
