# Quotes Scraper

A small Python scraper that collects quotes, authors, and tags from [Quotes to Scrape](https://quotes.toscrape.com/) and exports them to a formatted Excel workbook.

## Features

- Follows pagination and collects all quotes available on the site.
- Exports quote text, author, and tags to an `.xlsx` file.
- Formats the workbook with a filterable table, styled headers, readable column widths, and a frozen header row.
- Creates a timestamped workbook automatically if the default output file is open in Excel.

## Run

```bash
python -m pip install -r requirements.txt
python scraper.py
```

The workbook is saved beside the script as `extracted_data.xlsx`. Generated workbooks are excluded from Git; each user can create one locally by running the scraper.

## Data source

This project uses the public training website [Quotes to Scrape](https://quotes.toscrape.com/). The site provides a fixed sample dataset, so the scraper collects the available records rather than generating additional data.
