# Company Website & Email Scraper

An automated Python tool for finding and verifying company websites and extracting email addresses from an Excel file containing company names, addresses, and districts, currently focused on companies registered in Istanbul, Türkiye.

The tool uses Google Search to find potential company websites, evaluates candidates based on company-name similarity, verifies websites with Selenium, checks address/district information, extracts email addresses, and exports the results to Excel.

## Tech Stack

- Python 3
- Selenium
- Pandas
- OpenPyXL
- WebDriver Manager
- Google Search
- Web Scraping
- Data Extraction
- Automation

## Input

The program expects an Excel file named `firmalar.xlsx`.

The input file should contain at least these columns:

`UNVAN`, `ADRES`, `ILCE`

Example:

| UNVAN | ADRES | ILCE |
|---|---|---|
| Example Company Ltd. | Example Address | TUZLA |

## Output

The program creates `firmalar_web_mail.xlsx`.

The output contains:

`UNVAN`, `ITO_ADRES`, `ITO_ILCE`, `WEB`, `MAIL`, `WEB_ILCE`, `ADRES_DURUMU`, `DURUM`, `SITE_PUANI`

## Project Structure

```text
firmalar/
├── sitemailbul.py
├── firmalar.xlsx
├── firmalar_web_mail.xlsx
└── venv/
```

## Installation

Make sure Python 3 is installed.

Python 3.14.3 is currently used during development.

Clone or download this repository and navigate to the project directory.

Create a virtual environment:

- ### macOS

     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```

- ### Windows PowerShell

     ```bash
     python -m venv venv
     venv\Scripts\Activate
     ```

Install the required packages:

- ### macOS

     ```bash
     python3 -m pip install pandas openpyxl selenium webdriver-manager
     ```

- ### Windows

     ```bash
     python -m pip install pandas openpyxl selenium webdriver-manager
     ```

## Chrome Remote Debugging

The scraper connects to an existing Chrome session using Selenium remote debugging.

Chrome must be started with remote debugging enabled.

- ### macOS

  Open Terminal and run:

     ```bash
     "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --remote-debugging-port=9222 --user-data-dir="$HOME/chrome_selenium"
     ```

  Keep this Chrome window open while the scraper is running.

- ### Windows

  Open PowerShell and run:

     ```bash
     & "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="$env:USERPROFILE\chrome_selenium"
     ```

If Chrome is installed in a different location, update the path accordingly.

## Running the Program

Navigate to the project directory first.

- ### macOS

    ```bash
     cd ~/Desktop/firmalar
     source venv/bin/activate
     python3 sitemailbul.py
     ```
  
- ### Windows

    ```bash
     cd C:\firmalar
     venv\Scripts\Activate
     python sitemailbul.py
     ```

The project path may be different depending on where the repository is located.

> `firmalar.xlsx`, `firmalar_web_mail.xlsx`, and `venv/` contain local data or environment files and should not be committed to the public repository.
