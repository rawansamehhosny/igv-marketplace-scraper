# IGV Marketplace Enterprise Scraper and Automation Pipeline

A production-ready, resilient web scraping and data extraction pipeline engineered for IGV.com. The system extracts structured marketplace listings across gaming categories (Accounts, Items, Coins, and Boosting), enriches entries with seller store analytics and detailed descriptions, and enforces data integrity through automated checkpoints and a comprehensive QA test suite.

---

## Architectural Overview

The pipeline is structured into four functional layers:

1. Directory Discovery: Catalogs all available games, categories, and marketplace URLs.
2. Batch Partitioning: Segments large game inventories into manageable execution units.
3. Multi-Category Extraction Engine: Dynamically routes requests between heterogeneous API endpoints, enriches seller profiles, parses offer descriptions, and streams output.
4. Quality Assurance Suite: Runs end-to-end integration and data validation tests before production deployment.

---

## Key Engineering Features

### Dynamic API Endpoint Routing
IGV utilizes distinct internal APIs depending on the product classification:
- Accounts Marketplace: Routed to `/api/product-search/product-market/product/page` with specialized filtering flags, warranty periods, and guarantee tags.
- Standard Items and Currencies: Routed to `/api/product-search/v2/product/standard` utilizing the `mixPageResp` payload parser.
The scraper automatically detects the category and adjusts the HTTP method, payload structure, and response deserialization accordingly.

### Anti-Ban Mechanics and Rate Limit Auto-Recovery
- Adaptive Request Jitter: Implements randomized delay intervals (`polite_sleep`) centered around configurable thresholds to prevent fingerprinting by web application firewalls.
- Dynamic HTTP 429 Recovery: Evaluates incoming `Retry-After` HTTP response headers during rate-limit events, pausing execution for the requested duration before resuming with exponential backoff.
- Cloudflare and WAF 403 Mitigation: Catches forbidden status codes and applies stepped cooldown intervals before attempting retry attempts.
- Centralized Session Management: Dispatches all traffic via a hardened `safe_request` wrapper preventing unhandled network dropouts.

### Fault-Tolerant Batching and Incremental Checkpoints
- Real-Time Checkpointing: Flushes parsed records to disk after every retrieved page in UTF-8-SIG encoded CSV format, guarding against data loss during unexpected connection interruptions.
- Resume Engine (`AUTO_RESUME`): Inspects existing checkpoint files at launch and automatically skips games and categories that have already been collected.
- Batch Processing: Splits catalog datasets into isolated batches (default: 50 games per batch) allowing distributed scraping sessions or overnight executions.

### Comprehensive Entity Enrichment
- Seller Analytics: Resolves individual seller store profiles via dedicated API requests, capturing star rating level, operational tier, lifetime order volume, positive review counts, satisfaction percentage, KYC status, and delivery speed score.
- Tenure Calculation: Translates seller registration timestamps into formatted tenure intervals (e.g., `2 year(s) 3 month(s)`).
- Clean Description Parsing: Submits secondary POST requests to fetch full HTML descriptions, stripping tags and formatting while isolating embedded screenshot URLs.
- Specification Aggregation: Combines multiple variant attributes into a consolidated schema adhering to the format `Property//Value1--Value2`.

### Strict Data Sanitization
- Zero-NaN Guarantee: All missing, null, or unassigned data fields are scrubbed and replaced with deterministic values (`N/A` or empty strings for dynamic specifications).
- Price Standardization: Normalizes prices to USD formatting, eliminating trailing floating-point artifacts.

---

## Automated QA Test Suite

The repository includes a 12-step validation harness (`test_igv_pipeline.py`) structured across three testing phases:

### Phase 1: Directory and Categories Mapping
- TC-DIR-01: Verifies discovery of the Accounts category with greater than 100 mapped games.
- TC-DIR-02: Verifies discovery of the Items category with greater than 100 mapped games.
- TC-DIR-03: Validates platform-wide category discovery across all primary game markets.

### Phase 2: Pipeline Loop and Dynamic Routing
- TC-LOOP-01: Verifies endpoint switching to `/product-market/product/page` when processing Accounts.
- TC-LOOP-02: Verifies endpoint switching to `/v2/product/standard` when processing Items.
- TC-LOOP-03: Asserts dataset contains records representing both Accounts and Items.
- TC-LOOP-04: Asserts that each row in the resulting dataset is accurately tagged with its respective Category.

### Phase 3: Data Integrity and Fault Recovery
- TC-DATA-01: Enforces zero pandas `NaN` occurrences across all critical and dynamic specification columns.
- TC-DATA-02: Unit-tests attribute aggregation logic to guarantee `Property//Value1--Value2` formatting.
- TC-DATA-03: Validates that live scraped specifications adhere to the structured delimiter syntax.
- TC-DATA-04: Confirms the presence, readability, and row-count consistency of the checkpoint CSV.
- TC-ERR-01: Simulates HTTP 429 status code via unittest mocks to verify automatic `Retry-After` delay calculation and successful retry execution.

---

## Repository Structure

```text
.
├── .gitignore                          # Excludes data files, virtual environments, and caches
├── requirements.txt                    # Project dependencies
├── README.md                           # Technical documentation
├── get_all_games.py                    # Accounts catalog discovery script (570+ games)
├── build_all_categories_directory.py   # Multi-category platform indexing script
├── split_games_into_batches.py         # Batch generation utility for segmented runs
├── igv_marketplace_scraper.py          # Single-game intensive scraper (Fortnite / Accounts)
├── run_multi_game_sample.py            # Multi-game, multi-category scraping engine
└── test_igv_pipeline.py                # 12-step automated QA verification test suite
```

---

## Prerequisites and Installation

### System Requirements
- Python 3.9 or higher
- Windows, macOS, or Linux

### Environment Setup

1. Clone or download the repository:
```bash
git clone <REPOSITORY_URL>
cd igv-scraper
```

2. Create and activate a virtual environment:
```bash
# On Linux/macOS
python3 -m venv venv
source venv/bin/activate

# On Windows
python -m venv venv
venv\Scripts\activate
```

3. Install required dependencies:
```bash
pip install -r requirements.txt
```

---

## Usage Guide

### 1. Catalog Discovery
To extract the complete index of games offering Accounts:
```bash
python get_all_games.py
```
This generates `igv_games_list.xlsx` containing games, available offer counts, and direct category URLs.

To generate a cross-category directory covering Accounts, Items, Coins, and Boosting:
```bash
python build_all_categories_directory.py
```

### 2. Batch Partitioning
To partition the catalog into batch files of 50 games each:
```bash
python split_games_into_batches.py
```
Batch files will be placed in the `game_batches/` directory.

### 3. Running the Multi-Game Scraper
To run the multi-game scraper across configured batches or samples:
```bash
python run_multi_game_sample.py
```
Output data is saved incrementally to `igv_offers_batch_X_checkpoint.csv` and compiled into `igv_offers_batch_X.xlsx` upon completion.

### 4. Running a Focused Game Pilot
To run a deep scrape for a single game (e.g., Fortnite Accounts up to 50 pages):
```bash
python igv_marketplace_scraper.py
```

### 5. Running the QA Automation Suite
To validate the entire scraping pipeline, verify dynamic endpoint switching, and test anti-ban resilience:
```bash
python test_igv_pipeline.py
```

---

## Configuration Reference

Key execution settings can be adjusted directly in `run_multi_game_sample.py`:

| Parameter | Default | Description |
|---|---|---|
| `REQUEST_DELAY_SECONDS` | `1.5` | Base delay between page extraction requests |
| `ITEM_REQUEST_DELAY` | `0.5` | Delay between item-level seller and description queries |
| `BATCH_SIZE` | `50` | Number of games processed per batch run |
| `BATCH_NUMBER` | `1` | Target batch index (1 = games 1-50; 0 = process all) |
| `AUTO_RESUME` | `True` | Automatically bypasses targets found in checkpoint files |
| `PAGES_PER_GAME` | `1` | Page depth limit per target game |
| `MAX_RETRIES` | `3` | Maximum HTTP retry attempts on connection or server error |
| `TIMEOUT_SECONDS` | `25` | Maximum socket timeout duration per request |

---

## Data Schema

The scraper exports data with the following standardized fields:

| Field Name | Type | Description |
|---|---|---|
| `Game Name` | String | Standardized title of the game |
| `Category` | String | Target category (`Accounts`, `Items`, etc.) |
| `Offer Type` | String | Transaction type (`Sale`, `Rent`) |
| `Title` | String | Title of the marketplace listing |
| `Price` | String | Numeric price formatted to USD |
| `Currency` | String | Standard currency code (`USD`) |
| `Delivery Speed` | String | Delivery duration guarantee (`Instant`, `10 Min`, `Manual`) |
| `Warranty` | String | Seller warranty policy (e.g., `7d Warranty`, `Guaranteed`) |
| `Seller ID` | String | Unique platform identifier for the seller |
| `Seller Name` | String | Store name or username of the seller |
| `Seller Level` | String | Platform seller star level designation |
| `Seller Tier` | String | Operational tier name |
| `Seller Tenure` | String | Formatted time duration since registration |
| `Seller Member Since` | String | Registration date |
| `Seller Review Count` | String | Total reviews received |
| `Seller Positive Reviews`| String | Count of positive feedback ratings |
| `Seller Avg Rating` | String | Average store review score |
| `Seller Satisfaction Rate` | String | Percentage satisfaction rating |
| `Seller Completed Orders` | String | Total fulfilled transaction count |
| `Seller Sales Volume` | String | Volume metrics |
| `Seller Verified (KYC)` | String | Identity verification status (`Yes`, `No`) |
| `Seller Delivery Speed` | String | Seller delivery performance index |
| `Publish Date` | String | Listing publication timestamp |
| `Description` | String | Sanitized plain text description with image URLs |
| `Offer ID` | String | Unique marketplace listing ID |
| `Offer URL` | String | Direct canonical URL to the listing |
| `Images` | String | Pipe-delimited list of listing image URLs |
| `Page` | Integer | Source pagination index |
| `Item Index` | Integer | Position within the source page |
| `specification1` - `15+`| String | Delimited attributes in `Property//Value1--Value2` format |

---

## License

This project is developed for educational, analytical, and data evaluation purposes. Ensure compliance with target site terms of service when executing automated scrapers.
