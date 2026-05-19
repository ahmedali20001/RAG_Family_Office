# Extraction & Validation Methodology

This document outlines the pipeline used to build the verified dataset of Family Office investment advisers from the SEC IAPD portal.

## 1. How We Found Them
We utilized a seed list of 50 CRD (Central Registration Depository) numbers provided in `crd.csv`. This was checked and manually downloaded the list and trim it to 50. 
Using Python's `requests` library, we performed targeted HTTP GET requests to the hidden SEC IAPD REST API endpoints (e.g., `https://api.adviserinfo.sec.gov/IAPD/Content/Search/iapd_Search.aspx`). By passing the CRD numbers directly into the API payload, we bypassed the manual web interface and successfully downloaded the complete, raw JSON profiles for all 50 individuals.

## 2. How We Enriched Them
The raw SEC JSON is highly fragmented and deeply nested. We built a custom Python parser (`extract_family_offices.py`) to synthesize and enrich this data into a flat, analyzable structure:
*   **Experience Calculation:** We extracted the `daysInIndustryCalculatedDateIAPD` string, parsed the exact date, and dynamically calculated the "Years of Experience" integer to the present day.
*   **Geographic Aggregation:** We looped through both `currentEmployments` and `currentIAEmployments` arrays to deduplicate and concatenate Branch Office Locations into a clean "City, State" list.
*   **Disclosure Flattening:** Disciplinary and regulatory disclosures are nested incredibly deep (Allegations > Resolution > Sanctions > Fine Amounts). We wrote algorithms to traverse these arrays and aggregate them into two distinct, highly readable columns: `Total Disclosures` (integer) and `Disclosure Details` (a clean, human-readable paragraph containing the fine amounts and allegations).

## 3. How We Validated Them
To ensure the integrity of the data before feeding it to the AI, we implemented a deterministic Validation Engine that scores every profile out of 100%. 
A profile receives points if critical fields successfully parse:
*   **+20%**: Valid First/Last Name exists.
*   **+20%**: Valid IA (Investment Adviser) or BD (Broker/Dealer) scope string.
*   **+20%**: Industry Start Date is properly formatted and successfully parsed.
*   **+20%**: At least one active Branch Office Location is mapped.
*   **+20%**: The Disclosures array exists (even if 0, ensuring the API node didn't fail).

All 50 records were passed through this engine, and their individual confidence scores (e.g., 100%) were logged directly into a dedicated "Validation Chain" sheet inside the final Excel workbook to guarantee audit readiness.

## 4. What I Would Improve
If tasked with scaling this to 10,000+ Family Offices, I would implement the following improvements:
1.  **Asynchronous Scraping:** Replace the synchronous `time.sleep()` loop with `asyncio` and `aiohttp` to dramatically increase extraction speed.
2.  **Proxy Rotation & Rate Limiting:** The SEC actively blocks high-volume IPs. I would integrate a residential proxy pool to distribute the requests and avoid IP bans.
3.  **PDF Parsing (ADV Part 2A):** The JSON API only provides demographic and regulatory data. I would build a Selenium/Playwright scraper to download the firm's actual "Form ADV Part 2A" brochure (which is a PDF) and use OCR to extract their specific investment strategies and asset classes.
