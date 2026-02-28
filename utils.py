"""
Utility functions for the Familienpass Event Scraper
"""

import time
import re
import requests
from typing import Optional
from bs4 import BeautifulSoup, Tag
from config import USER_AGENT, REQUEST_TIMEOUT, MAX_RETRIES, INITIAL_BACKOFF


def make_request_with_retry(url: str, max_retries: int = MAX_RETRIES,
                            delay: int = INITIAL_BACKOFF) -> requests.Response:
    """
    Make HTTP request with retry logic and exponential backoff

    Args:
        url: URL to fetch
        max_retries: Maximum number of retry attempts
        delay: Initial delay between retries (doubles on each retry)

    Returns:
        Response object

    Raises:
        requests.exceptions.RequestException: If all retries fail
    """
    for attempt in range(max_retries):
        try:
            response = requests.get(
                url,
                headers={'User-Agent': USER_AGENT},
                timeout=REQUEST_TIMEOUT
            )
            response.raise_for_status()
            return response

        except requests.exceptions.Timeout:
            print(f"  Timeout on attempt {attempt + 1}/{max_retries}")
            if attempt < max_retries - 1:
                time.sleep(delay * (2 ** attempt))  # Exponential backoff
            else:
                raise

        except requests.exceptions.HTTPError as e:
            if e.response.status_code in [429, 503]:  # Rate limit or service unavailable
                print(f"  Server busy (HTTP {e.response.status_code}), waiting...")
                time.sleep(delay * 2)
                if attempt < max_retries - 1:
                    continue
            raise

        except requests.exceptions.RequestException as e:
            print(f"  Network error: {e}")
            if attempt < max_retries - 1:
                time.sleep(delay)
            else:
                raise

    raise requests.exceptions.RequestException(f"Failed to fetch {url} after {max_retries} attempts")


def clean_text(text: str) -> str:
    """
    Clean extracted text by removing extra whitespace and newlines

    Args:
        text: Raw text to clean

    Returns:
        Cleaned text
    """
    if not text:
        return ''

    # Remove extra whitespace and newlines
    text = re.sub(r'\s+', ' ', text)

    # Strip leading/trailing whitespace
    text = text.strip()

    return text


def extract_field_by_header(soup: BeautifulSoup, header_text: str) -> str:
    """
    Find <h3> header with specific text and extract following content

    Args:
        soup: BeautifulSoup object of the page
        header_text: Text to search for in <h3> headers

    Returns:
        Cleaned text content following the header, or empty string if not found
    """
    try:
        # Find all h3 headers
        headers = soup.find_all('h3')

        for header in headers:
            if header_text.lower() in header.get_text().lower():
                # Get next sibling content
                next_elem = header.find_next_sibling()

                if next_elem:
                    return clean_text(next_elem.get_text())

                # Alternative: get all text until next header
                content = []
                for sibling in header.next_siblings:
                    if sibling.name == 'h3':
                        break
                    if hasattr(sibling, 'get_text'):
                        content.append(sibling.get_text())

                if content:
                    return clean_text(' '.join(content))

        return ''

    except Exception as e:
        print(f"    Warning: Could not extract {header_text}: {e}")
        return ''
