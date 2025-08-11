# Company Name Availability Checker

This Python script automates the process of checking company name availability on the Czech Business Register ([justice.cz](https://or.justice.cz/)) by reading a list of names from an Excel file, performing searches via Selenium, and writing the results back to the same file.

## Features
- Reads names from an Excel sheet (`names.xlsx`).
- Searches each name on the Czech Business Register website.
- Records the number of matching entries found.
- Marks names with no matches as **`free`**.
- Sorts results so available names appear at the top.
- Saves results back into the original Excel file.

---

## Requirements

### Python
- Python 3.8 or higher

### Dependencies
Install required packages:
```bash
pip install pandas selenium openpyxl webdriver-manager
