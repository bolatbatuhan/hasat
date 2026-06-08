#  Hasat — Web Wordlist & OSINT Harvester

Hasat is a modular, high-performance **web wordlist generator and OSINT** tool engineered specifically for penetration testers and Red Team operators during engagement reconnaissance phases.

Unlike standard wordlist generators, Hasat goes beyond flat HTML parsing. It renders dynamic JavaScript single-page applications (SPAs), extracts hidden context from file attachments (PDF, Word, Excel), grabs target email addresses, and automatically derives customized user brute-force lists based on corporate naming conventions.

---

##  Core Features

* **Recursive Deep Crawling:** Automatically maps internal architecture based on the specified `--depth` threshold.
* **Dynamic JavaScript Rendering (`--js`):** Leverages Playwright integration to execute client-side JavaScript, ensuring no words are missed on React, Angular, or Vue-based applications.
* **Document Content Extraction (`--files`):** Downloads and parses `.pdf`, `.docx`, and `.xlsx` files in the background to scrape hidden corporate terminology, policies, and internal keywords.
* **Advanced OSINT Engine (`--email`):** Extracts raw email addresses and translates standard schemas (`firstname.lastname@target.local`, etc.) into structured brute-force username dictionaries.
* **Hidden Metadata Scraper:** Tracks HTML `meta` tags and image `alt` attributes to unearth contextually relevant keywords invisible to casual site visitors.
* **WAF / IPS Evasion Controls:** Throttles traffic via `--delay`, randomizes footprints with custom `User-Agent` masks, routes requests through proxies (e.g., Burp Suite via `--proxy`), and appends authentication headers or session cookies (`--header`).
* **Fault-Tolerant Modular Design:** Missing third-party packages (like Playwright or pdfplumber) won't cause execution crashes. The script dynamically catches exceptions and degrades gracefully to core HTML scraping capabilities.

---

##  Installation

Hasat requires Python 3.8+ environments.

Clone the repository to your local machine:
```bash
git clone https://github.com/bolatbatuhan/hasat.git
cd hasat

Install all necessary and optional dependencies:

pip install -r requirements.txt

(Optional) Install the web engine binaries to unlock full JavaScript rendering support:

playwright install chromium
```

##  CLI Arguments Reference
hasat [-h] -u URL [-d DEPTH] [-o OUTPUT] [-v] [--delay DELAY] [--min-word MIN_WORD]
      [--max-word MAX_WORD] [--js] [--files] [--email] [--auth-user AUTH_USER]
      [--auth-pass AUTH_PASS] [--auth-type {basic,digest}] [--user-agent USER_AGENT]
      [--header HEADER] [--proxy PROXY] [--no-verify-ssl]

##  Argument Breakdown
| Argument | Description | Default |
| :--- | :--- | :--- |
| `-u, --url` | Target web application root URL (Required) | — |
| `-d, --depth` | Recursive link crawling boundary depth limit | `2` |
| `-o, --output` | Filename path for the generated unique wordlist output | `wordlist.txt` |
| `-v, --verbose` | Enable real-time logging output for live intelligence monitoring | `False` |
| `--delay` | Interval pause between server requests to control rate limits (Seconds) | `0.5` |
| `--min-word` | Minimum character string length constraint for word preservation | `4` |
| `--max-word` | Maximum character string length constraint for filtering long blobs | `25` |
| `--js` | Activates headless browser actions via Playwright for dynamic page execution | `False` |
| `--files` | Activates document scraping processors for PDF, DOCX, and XLSX files | `False` |
| `--email` | Runs the regex intelligence engines for email harvesting and user derivation | `False` |
| `--proxy` | Configures upstream proxies (e.g., Burp Suite via http://127.0.0.1:8080) | — |
| `--no-verify-ssl` | Drops validation checks for insecure or self-signed internal TLS certificates | `False` |

##  License & Legal Disclaimer
This utility is strictly developed for authorized penetration testing, defensive research, and educational exercises. Scanning web applications without explicit prior consent or scope agreements is illegal and holds high criminal liability. The developer assumes no responsibility for unauthorized malicious actions.

Distributed under the MIT License.
