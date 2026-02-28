# Münchner Familienpass Tools

Python tools to
- Extract Familienpass event data from the Munich Familienpass website and export it to a formatted Excel spreadsheet
- Create create calendar events (.ics file) from the Excel sheet
- Create Apple Reminders for sign-up deadlines from the Excel sheet

**Target website**: [Ferienangebote Familienpass München](https://veranstaltungen.muenchen.de/ferienangebote-familienpass/familienpassangebote/)

## Features

- Scrapes all listing pages (~100+ events depending on the season)
- Extracts: event name, description, age requirement, location, date, time, and sign-up period
- Groups events with multiple dates together
- Exports to a professionally formatted Excel (.xlsx) file with clickable event links
- Preserves your event selections when re-running the scraper
- Respectful rate limiting (1 s between pages, 0.5 s between events)
- Retry logic with exponential backoff for network errors
- Export sign-up deadlines to a `.ics` calendar file
- Create Apple Reminders for sign-up deadlines

## Requirements

- Python 3.8+
- macOS is required only for the Apple Reminders feature (`create_reminder.py`)

## Installation

### 1. Create a virtual environment (recommended)

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

> **Note**: `pyremindkit` and `pyobjc-framework-EventKit` are macOS-only packages used by `create_reminder.py`. On other platforms you can remove those two lines from `requirements.txt` if you don't need the Reminders feature.

## Usage

### Step 1: Scrape events

```bash
python scraper.py
```

The script will:

1. Scrape all listing pages
2. Extract details from each event page (~60–90 seconds total)
3. Save results to `output/familienpass_events.xlsx`

Re-running the scraper will preserve any selections you have made in column A of the Excel file.

### Step 2: Mark events of interest (for calendar / reminders)

Open `output/familienpass_events.xlsx` and put any value (e.g. `x`) in the **"Selected"** column (column A) next to the events you want reminders for. Save the file.

### Step 3a: Create a calendar file (optional)

```bash
python create_calendar.py
```

This creates `output/familienpass_calendar.ics` with all-day calendar events covering the sign-up period for each selected event. The event title is `"Anmeldung <event name>"`.

Import into your calendar app:

- **macOS Calendar**: double-click the `.ics` file, or File → Import
- **Google Calendar**: Settings → Import
- **Outlook**: File → Open & Export → Import/Export

Events where registration is handled directly by the organiser ("direkt beim Veranstalter") or that have no sign-up date are skipped.

### Step 3b: Create Apple Reminders (macOS only, optional)

```bash
python create_reminder.py
```

Use `--dry-run` to preview what would be created without actually adding reminders:

```bash
python create_reminder.py --dry-run
```

Reminders are added to a **"Familienpass"** list (created automatically if it doesn't exist). Each reminder includes:

- **Title**: `Anmeldung Familienpass: <event name>`
- **Due date**: Sign-up period start date
- **Notes**: Event URL, sign-up period, event date and time

> **macOS permission**: The first run may prompt you to grant Reminders access to your terminal. If it fails silently, go to System Settings → Privacy & Security → Reminders and enable your terminal app.

## Output format

The Excel file contains the following columns:

| Column | Content |
|--------|---------|
| A | Selected (your marker) |
| B | Event name (clickable hyperlink) |
| C | Description |
| D | Age requirement (Alter) |
| E | Location (Treffpunkt) |
| F | Date |
| G | Time |
| H | Sign-up period (Verlosungszeitraum) |

Formatting: bold header row (dark blue / white), frozen header, 150 % zoom, text wrapping, top alignment.

## Configuration

Edit `config.py` to adjust:

| Setting | Default | Description |
|---------|---------|-------------|
| `TOTAL_PAGES` | `4` | Number of listing pages to scrape |
| `DELAY_BETWEEN_PAGES` | `1.0` | Seconds between listing pages |
| `DELAY_BETWEEN_EVENTS` | `0.5` | Seconds between event detail pages |
| `REQUEST_TIMEOUT` | `10` | HTTP request timeout in seconds |
| `MAX_RETRIES` | `3` | Maximum retry attempts on network errors |

## Project structure

```
Familienpass/
├── scraper.py          # Main scraping script
├── create_calendar.py  # Calendar (.ics) creator
├── create_reminder.py  # Apple Reminders creator (macOS only)
├── utils.py            # Shared helper functions
├── config.py           # Configuration constants
├── requirements.txt    # Python dependencies
├── LICENSE
├── README.md
└── output/             # Generated files (gitignored)
    ├── familienpass_events.xlsx
    └── familienpass_calendar.ics
```

## Troubleshooting

**"No module named ..."** — install dependencies: `pip install -r requirements.txt`

**No events found** — check your internet connection and verify the website is accessible.

**Scraper is slow** — this is intentional (rate limiting). You can lower the delays in `config.py`, but please be respectful to the server.

**Calendar shows no events** — make sure you marked events in column A and saved the Excel file before running `create_calendar.py`.

**Reminders permission error** — grant Reminders access to your terminal in System Settings → Privacy & Security → Reminders.

## Notes

- **Data accuracy**: Always verify critical information on the original website before registering.
- **Educational / personal use**: This scraper is intended for personal use. Please respect the website's terms of service.
- **Website changes**: If the scraper stops working, the website HTML structure may have changed. Check the `<h3>` headers and table structure on the listing pages.

## License

MIT — see [LICENSE](LICENSE).
