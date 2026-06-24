# Persona Report Generator MVP

Excel upload based persona report generator.

## Features

- Upload `.xlsx` or `.xls` files in the browser
- Preview the first 10 rows
- Select columns for persona analysis
- Generate up to 3 rule-based persona cards
- View charts from uploaded Excel data
- Choose chart types: bar, line, pie, doughnut, polar area, radar
- Download the persona report as PDF

## Run

Open `index.html` in a browser.

For local testing with a simple server:

```bash
python -m http.server 8000
```

Then open `http://127.0.0.1:8000/`.
