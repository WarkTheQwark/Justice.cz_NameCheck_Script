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
```

### Browser Driver
- Google Chrome is required.
- The script uses webdriver-manager to automatically handle the correct ChromeDriver installation.

## Input File Format

The script expects a file named `names.xlsx` in the same directory as the script.

### Excel file requirements:
- Sheet name: `Sheet1` — Default Excel option
- Column: `NAMES` (case-sensitive) containing the list of names to check.
- Optional column: `MATCH_COUNT` — will be created automatically if missing.

### Example
| NAMES         | MATCH_COUNT |
|---------------|-------------|
| Example One   |             |
| Another Name  |             |

## Usage
1. Prepare your Excel file
- Place names.xlsx in the same directory as check_names.py.
2. Run the script
   ```bash
   python check_names.py
   ```
3. View results
The script updates names.xlsx:
- MATCH_COUNT = free → name is available.
- MATCH_COUNT = <number> → number of matching entries found.
- Available names are listed first.

## Notes
Default behavior launches Chrome in non-headless mode so you can see the searches.
To run in headless mode, uncomment:
`options.add_argument("--headless")`

You may need to increase the time.sleep() delays if your internet connection is slow, as too short a delay can cause missed results.

The script specifically targets the default Czech Business Register search form at:
https://or.justice.cz/ias/ui/rejstrik-$firma

If you wish to use different settings, make one manual serach with desired settings and use the resulting URL instead, justice.cz will references settings in the URL and is consistent between additional requests.
