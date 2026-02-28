"""
Configuration constants for Familienpass Event Scraper
"""

# Base URL for the event listing pages
BASE_URL = 'https://veranstaltungen.muenchen.de/ferienangebote-familienpass/familienpassangebote/'

# Number of pages to scrape
TOTAL_PAGES = 4

# Rate limiting settings (in seconds)
DELAY_BETWEEN_PAGES = 1.0       # Delay between listing pages
DELAY_BETWEEN_EVENTS = 0.5      # Delay between individual event pages

# HTTP request settings
REQUEST_TIMEOUT = 10            # Request timeout in seconds
MAX_RETRIES = 3                 # Maximum number of retry attempts
INITIAL_BACKOFF = 2             # Initial backoff delay for retries (in seconds)

# User-Agent header to identify the scraper
USER_AGENT = 'Mozilla/5.0 (compatible; FamilienpassScraper/1.0; Educational Purpose)'

# Marker used in the Event Name column for continuation rows (multi-date events)
CONTINUATION_LINK_MARKER = '↗'
