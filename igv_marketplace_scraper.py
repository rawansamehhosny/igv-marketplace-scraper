"""IGV Marketplace Scraper - Fortnite Accounts
Configured with verified browser headers, payload, enriched seller profiles,
full offer descriptions, seller tenure duration, and explicit offer types.
"""

from __future__ import annotations

from datetime import datetime
import json
import random
import re
import sys
import time
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

from bs4 import BeautifulSoup
import pandas as pd
import requests

# Fix Windows console UTF-8 display
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ---------- CONFIGURATION ----------
ENDPOINT = "https://www.igv.com/api/product-search/product-market/product/page"
SELLER_STORE_ENDPOINT = "https://www.igv.com/api/operation/seller/store"
PRODUCT_DETAIL_ENDPOINT = "https://www.igv.com/api/product-search/v2/product/standard/"
SITE_BASE_URL = "https://www.igv.com/"
REQUEST_PAYLOAD_FILE = Path("request_payload.json")
OUTPUT_XLSX = Path("igv_offers_pilot.xlsx")
OUTPUT_CSV = Path("igv_offers_checkpoint.csv")
DEBUG_RESPONSE_FILE = Path("debug_response_page1.json")

# Game & Category Info
GAME_NAME = "Fortnite"
CATEGORY_NAME = "Accounts"
GAME_CATEGORY_URL = "https://www.igv.com/fortnite/accounts"
TENANT_ID = "3332001"
SITE_ID = "1628721356376162310"

# Set to 50 pages for large scale stress test & rate-limit discovery
MAX_PAGES: int | None = 50
REQUEST_DELAY_SECONDS = 1.5
TIMEOUT_SECONDS = 30
MAX_RETRIES = 3

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "lang": "en",
    "timezone": "Africa/Cairo",
    "deviceid": "51a2d977eca7ac4e6802995744a18812",
    "ab-test-group": "1",
    "tenant-id": TENANT_ID,
    "referer": GAME_CATEGORY_URL,
    "origin": "https://www.igv.com",
}

# Cache for seller store profile to prevent duplicate HTTP requests
seller_cache: dict[str, dict[str, Any]] = {}
ITEM_REQUEST_DELAY = 0.5  # Polite delay between item-level requests to prevent 429/403

def polite_sleep(base_seconds: float, jitter_ratio: float = 0.25) -> None:
    """Sleeps with random jitter around base_seconds to avoid bot detection patterns."""
    low = max(0.1, base_seconds * (1.0 - jitter_ratio))
    high = base_seconds * (1.0 + jitter_ratio)
    time.sleep(random.uniform(low, high))

# Real-time request monitoring for Rate Limiting analysis
REQUEST_STATS = {
    "total_requests": 0,
    "rate_limit_429_hits": 0,
    "first_rate_limit_hit": None,
    "server_error_5xx_hits": 0,
    "rate_limit_events": [],
}


def safe_request(
    session: requests.Session,
    method: str,
    url: str,
    *,
    params: dict[str, Any] | None = None,
    json_data: dict[str, Any] | None = None,
    headers: dict[str, str] | None = None,
    timeout: int = TIMEOUT_SECONDS,
    max_retries: int = MAX_RETRIES,
) -> requests.Response | None:
    """
    Centralized HTTP request handler with automatic retry for 429 (respecting Retry-After),
    5xx server errors, and connection glitches with exponential backoff.
    """
    for attempt in range(max_retries + 1):
        REQUEST_STATS["total_requests"] += 1
        try:
            resp = session.request(
                method,
                url,
                params=params,
                json=json_data,
                headers=headers,
                timeout=timeout,
            )
            # Handle rate limiting (429)
            if resp.status_code == 429:
                REQUEST_STATS["rate_limit_429_hits"] += 1
                retry_after_hdr = resp.headers.get("Retry-After")
                if REQUEST_STATS["first_rate_limit_hit"] is None:
                    REQUEST_STATS["first_rate_limit_hit"] = {
                        "request_number": REQUEST_STATS["total_requests"],
                        "url": url,
                        "time": datetime.now().strftime("%H:%M:%S"),
                        "retry_after": retry_after_hdr,
                    }
                event_info = f"Req #{REQUEST_STATS['total_requests']} | URL: {url.split('?')[0]} | Retry-After: {retry_after_hdr}"
                REQUEST_STATS["rate_limit_events"].append(event_info)

                if retry_after_hdr and retry_after_hdr.strip().isdigit():
                    wait_time = float(retry_after_hdr.strip())
                else:
                    wait_time = float(2 ** attempt + 1.5)
                print(f"\n  ⚠️ [RATE LIMIT 429 Triggered at Req #{REQUEST_STATS['total_requests']}] Waiting {wait_time:.1f}s before retrying: {url}")
                time.sleep(wait_time)
                continue

            # Handle Cloudflare / WAF 403 Forbidden
            if resp.status_code == 403:
                print(f"\n  🛑 [Cloudflare / WAF 403 Forbidden at Req #{REQUEST_STATS['total_requests']}] Access blocked: {url}")
                if attempt < max_retries:
                    wait_time = float(5.0 * (attempt + 1))
                    print(f"  🛑 Waiting {wait_time:.1f}s cooldown before retrying...")
                    time.sleep(wait_time)
                    continue
                return resp

            # Handle transient server errors (5xx)
            if resp.status_code in (500, 502, 503, 504):
                REQUEST_STATS["server_error_5xx_hits"] += 1
                if attempt < max_retries:
                    wait_time = float(2 ** attempt)
                    print(f"  [Server Error {resp.status_code} at Req #{REQUEST_STATS['total_requests']}] Retry in {wait_time:.1f}s: {url}")
                    time.sleep(wait_time)
                    continue
                return resp

            return resp

        except requests.RequestException as e:
            if attempt >= max_retries:
                print(f"  [Network Error] Final failure for {url}: {e}")
                return None
            wait_time = float(2 ** attempt)
            time.sleep(wait_time)

    return None


def load_payload() -> dict[str, Any]:
    payload: dict[str, Any] = {}
    if REQUEST_PAYLOAD_FILE.exists():
        try:
            data = json.loads(REQUEST_PAYLOAD_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                payload = data
        except Exception as e:
            print(f"Warning reading {REQUEST_PAYLOAD_FILE}: {e}")

    # Fallback default verified payload if file missing or empty
    if not payload:
        payload = {
            "tenantId": TENANT_ID,
            "isTenantOnSale": True,
            "pageIndex": 1,
            "pageSize": 16,
            "brandIds": ["2728"],
            "isShowAttrTag": False,
            "isShowAttr": True,
            "isShowGuarantee": True,
            "productCategoryAttrs": [],
            "keywords": "",
            "tenantCategoryId3": "76498313579",
            "automationKeywords": "",
            "sorts": [{"field": "bestMatchScore", "isAsc": False}],
            "isShowSellerInfo": True,
            "isShowSellerStar": True,
            "siteId": SITE_ID,
            "toCurrencyId": "1598942409803481063",
        }

    # ضمان إرسال معرّف عملة الدولار (USD) حصراً في الطلب
    payload["toCurrencyId"] = "1598942409803481063"
    return payload


def calculate_seller_tenure(register_time_str: str) -> str:
    """
    Calculates the exact duration a seller has been on the platform (e.g. '2 year(s) 3 month(s)').
    """
    if not register_time_str or str(register_time_str).strip() in ("", "N/A", "None"):
        return "N/A"
    try:
        dt = datetime.strptime(str(register_time_str).split()[0], "%Y-%m-%d")
        now = datetime.now()
        days = (now - dt).days
        if days < 0:
            return "Recently joined"
        years = days // 365
        remaining_days = days % 365
        months = remaining_days // 30
        if years >= 1:
            if months > 0:
                return f"{years} year(s) {months} month(s)"
            return f"{years} year(s)"
        elif months >= 1:
            return f"{months} month(s)"
        else:
            return f"{days} day(s)"
    except Exception:
        return "N/A"


def fetch_seller_profile(session: requests.Session, seller_id: str) -> tuple[dict[str, Any], bool]:
    """
    Fetches full store profile for a seller (registerTime, reviewCount, averageScore, etc.).
    Uses local cache to avoid duplicate requests. Only caches successful responses.
    Returns: (seller_info_dict, is_success)
    """
    if not seller_id or str(seller_id).strip() in ("", "None", "N/A"):
        return {}, True

    seller_id_str = str(seller_id).strip()
    if seller_id_str in seller_cache:
        return seller_cache[seller_id_str], True

    params = {
        "sellerId": seller_id_str,
        "tenantId": TENANT_ID,
        "siteId": SITE_ID,
    }
    seller_headers = dict(HEADERS)
    seller_headers["referer"] = f"https://www.igv.com/store/{seller_id_str}"

    polite_sleep(ITEM_REQUEST_DELAY)
    resp = safe_request(session, "GET", SELLER_STORE_ENDPOINT, params=params, headers=seller_headers, timeout=10)
    if resp is not None and resp.status_code == 200:
        try:
            data = resp.json()
            if data.get("code") in ("0000", "0", "200"):
                info = data.get("data") or {}
                if isinstance(info, dict) and info:
                    seller_cache[seller_id_str] = info
                    return info, True
                return {}, True
        except Exception as e:
            print(f"Warning parsing seller {seller_id_str}: {e}")
            return {}, False

    # Note: Do NOT store empty dict in seller_cache on failure, so temporary failures can be retried
    return {}, False


def fetch_offer_description(session: requests.Session, offer_id: str) -> tuple[str, str]:
    """
    Fetches full offer description HTML from product detail API, extracting clean text
    and image URLs, never returning raw HTML.
    Returns: (cleaned_description, status) where status is 'ok', 'empty', or 'failed'.
    """
    if not offer_id or str(offer_id).strip() in ("", "None", "N/A"):
        return "N/A", "empty"

    url = f"{PRODUCT_DETAIL_ENDPOINT}{offer_id}"
    payload = {"siteId": SITE_ID, "tenantId": TENANT_ID}
    headers = dict(HEADERS)
    headers["referer"] = f"{GAME_CATEGORY_URL.rstrip('/')}/{offer_id}"

    polite_sleep(ITEM_REQUEST_DELAY)
    resp = safe_request(session, "POST", url, json_data=payload, headers=headers, timeout=10)
    if resp is not None and resp.status_code == 200:
        try:
            data = resp.json()
            if data.get("code") in ("0000", "0", "200"):
                raw_desc = (
                    data.get("data", {})
                    .get("noResp", {})
                    .get("description")
                )
                if raw_desc and str(raw_desc).strip():
                    soup = BeautifulSoup(str(raw_desc), "html.parser")
                    text = soup.get_text(separator="\n").strip()
                    img_tags = soup.find_all("img")
                    img_urls = [str(img.get("src")).strip() for img in img_tags if img.get("src")]

                    if text and img_urls:
                        return f"{text}\n[Description Images: {' | '.join(img_urls)}]", "ok"
                    elif text:
                        return text, "ok"
                    elif img_urls:
                        return f"[Image-based Description: {' | '.join(img_urls)}]", "ok"
                    else:
                        return "N/A", "empty"
                else:
                    return "N/A", "empty"
        except Exception as e:
            print(f"Warning parsing description for offer {offer_id}: {e}")
            return "N/A", "failed"

    return "N/A", "failed"


def _is_id_value(val: str) -> bool:
    """Returns True if val looks like an internal database or platform ID rather than human-readable text."""
    if not val or not isinstance(val, str):
        return True
    v = val.strip()
    if not v or v.lower() in ("none", "null", "undefined", "n/a", "na", ""):
        return True
    # 1. MongoDB ObjectId: 24 hex characters
    if len(v) == 24 and re.fullmatch(r"[0-9a-fA-F]{24}", v):
        return True
    # 2. UUID or 32-char hex hash
    if re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", v):
        return True
    if len(v) == 32 and re.fullmatch(r"[0-9a-fA-F]{32}", v):
        return True
    # 3. Raw long numeric platform ID (>= 8 digits)
    if v.isdigit() and len(v) >= 8:
        return True
    return False


def normalize_delivery_to_minutes(d_val: Any, d_unit: Any = None) -> str:
    """
    Normalizes delivery time to pure integer minutes without any words or letters.
    """
    if d_val is None:
        return "N/A"
    if d_unit is not None:
        try:
            val_num = int(float(str(d_val).strip()))
            unit_code = int(str(d_unit).strip())
            if unit_code == 3:  # Minutes
                return str(val_num)
            elif unit_code == 2:  # Hours
                return str(val_num * 60)
            elif unit_code == 1:  # Days
                return str(val_num * 1440)
        except (ValueError, TypeError):
            pass

    s = str(d_val).strip()
    if not s or s.lower() in ("none", "null", "n/a", "", "manual / specified", "manual", "unknown"):
        return "N/A"
    if s.lower() in ("instant", "0", "0 min", "0m"):
        return "0"

    m = re.search(r'(\d+(?:\.\d+)?)\s*(days?|d|hours?|hrs?|h|minutes?|mins?|m)?', s, re.IGNORECASE)
    if m:
        num_str, unit = m.group(1), m.group(2)
        try:
            num = float(num_str)
            if not unit:
                return str(int(num))
            u = unit.lower()
            if u.startswith("d"):
                return str(int(num * 1440))
            elif u.startswith("h"):
                return str(int(num * 60))
            elif u.startswith("m"):
                return str(int(num))
        except (ValueError, TypeError):
            pass
    return "N/A"


def extract_tags(item: dict[str, Any]) -> list[str]:
    attrs = item.get("productSpuCategoryAttrs") or []
    grouped_tags: dict[str, list[str]] = {}
    for attr in attrs:
        if isinstance(attr, dict):
            name = str(attr.get("categoryAttrName") or "").strip()
            val = str(attr.get("categoryAttrValueName") or "").strip()
            # Skip empty values or raw internal IDs
            if not val or _is_id_value(val):
                continue
            if name and _is_id_value(name):
                continue
            if name:
                grouped_tags.setdefault(name, []).append(val)
            else:
                grouped_tags.setdefault("Other", []).append(val)
    tags: list[str] = []
    for name, vals in grouped_tags.items():
        unique_vals = list(dict.fromkeys(vals))
        if name == "Other":
            tags.extend(unique_vals)
        else:
            tags.append(f"{name}//{'--'.join(unique_vals)}")
    return tags


def extract_images(item: dict[str, Any]) -> list[str]:
    """
    Extracts all product image URLs from productSpuImages and productImages.
    """
    images = item.get("productSpuImages") or item.get("productImages") or []
    img_list = []
    if isinstance(images, list):
        for img in images:
            if isinstance(img, str):
                img_list.append(img)
            elif isinstance(img, dict):
                u = img.get("imageUrl") or img.get("url") or img.get("src")
                if u:
                    img_list.append(str(u))
    return img_list


def parse_offer(
    item: dict[str, Any],
    page_index: int,
    item_index: int,
    seller_info: dict[str, Any] | None = None,
    description: str = "N/A"
) -> dict[str, Any]:
    if seller_info is None:
        seller_info = {}

    tags = extract_tags(item)

    # Helper to extract all matching tag values (case‑insensitive)
    def _extract_all_matching(prefix: str) -> list[str]:
        """Return a list of values for tags whose name starts with *prefix*.
        If a tag contains multiple values separated by '--', they are split.
        Returns empty list if none found."""
        values = []
        for tag in tags:
            if tag.lower().startswith(prefix.lower() + "//"):
                # tag format: Name//val1--val2
                vals_part = tag.split("//", 1)[1]
                for v in vals_part.split("--"):
                    v_clean = v.strip()
                    if v_clean:
                        values.append(v_clean)
        return values

    # 1. تحديد نوع الأوفر
    tx_type = item.get("transactionType")
    offer_type = "Sale" if tx_type == 1 else ("Rent" if tx_type == 2 else "N/A")

    # 2. استخراج التوصيل والضمان
    # بناءً على استجابة الـ API، الضمان بيكون بالأيام
    guarantee_days = item.get("categoryGuaranteeDays")
    warranty = f"{guarantee_days} Days" if guarantee_days else "No Warranty"

    # التوصيل: تحويل التوصيل إلى دقائق فقط بدون حروف
    delivery_code = item.get("deliveryType")
    if str(delivery_code).strip() == "1":
        delivery_time_min = "0"
    elif str(delivery_code).strip() == "2":
        delivery_time_min = "10"
    else:
        delivery_time_min = normalize_delivery_to_minutes(delivery_code)

    # 3. الصور والرابط
    img_list = extract_images(item)
    offer_id = str(item.get("id") or item.get("productSpuId") or "")
    detail_url = urljoin(SITE_BASE_URL, str(item.get("spuItemUrl") or f"{GAME_CATEGORY_URL.rstrip('/')}/{offer_id}"))

    # 4. بيانات البائع
    seller_name = (
        seller_info.get("storeName")
        or item.get("sellerStoreName")
        or item.get("sellerName")
        or "N/A"
    )

    raw_rate = seller_info.get("reviewRate")
    review_rate_str = "N/A"
    if raw_rate not in (None, ""):
        try:
            review_rate_str = f"{float(raw_rate) * 100:.1f}%"
        except (ValueError, TypeError):
            review_rate_str = str(raw_rate)

    is_kyc = seller_info.get("isCompletedEKyc")
    verified_str = "Yes" if is_kyc is True else ("No" if is_kyc is False else "N/A")

    reg_time = seller_info.get("registerTime") or "N/A"
    seller_tenure = calculate_seller_tenure(reg_time)
    
    # 5. السعر
    skus = item.get("productSkus") or []
    price = "N/A"
    for sku in skus:
        if isinstance(sku, dict) and (str(sku.get("currencyCode")).upper() == "USD" or str(sku.get("currencyId")) == "1598942409803481063"):
            price = sku.get("price")
            break
    if price in ("N/A", None):
        price = item.get("price") or item.get("salePrice") or "N/A"

    # تنظيف وتنسيق السعر لتجنب الكسور الزائدة
    if price not in ("N/A", None, ""):
        try:
            val = float(price)
            price = f"{val:.2f}".rstrip("0").rstrip(".") if "." in f"{val:.2f}" else f"{val:.0f}"
        except (ValueError, TypeError):
            pass

    # 6. تجميع الصف
    row: dict[str, Any] = {
        "Game Name": GAME_NAME,
        "Category": CATEGORY_NAME,
        "Offer Type": offer_type,  # عمود مخصص لنوع الأوفر (Sale / Rent / N/A)
        "Title": item.get("title") or item.get("productSpuName") or "N/A",
        "Price": price,
        "Currency": "USD",
        "Delivery Time (Min)": str(delivery_time_min),
        "Warranty": warranty,                # <--- العمود الجديد
        "Seller ID": item.get("sellerId") or "N/A",
        "Seller Name": seller_name,
        "Seller Level": seller_info.get("sellerStarName") or item.get("sellerStarName") or "N/A",
        "Seller Tier": seller_info.get("tierName") or "N/A",
        "Seller Tenure": seller_tenure,  # مدة عمل السيلر (سنة / سنتين ...)
        "Seller Member Since": reg_time,
        "Seller Review Count": seller_info.get("reviewCount") or "N/A",
        "Seller Positive Reviews": seller_info.get("niceReviewCount") or "N/A",
        "Seller Avg Rating": seller_info.get("averageScore") or "N/A",
        "Seller Satisfaction Rate": review_rate_str,
        "Seller Completed Orders": seller_info.get("orderCount") or "N/A",
        "Seller Sales Volume": seller_info.get("salesVolume") or "N/A",
        "Seller Verified (KYC)": verified_str,
        "Seller Delivery Speed": seller_info.get("deliverySpeedScore") or "N/A",
        "Publish Date": item.get("publishDate") or item.get("createTime") or "N/A",
        "Description": description,  # الوصف الكامل للعرض بدون HTML
        "Offer ID": offer_id or "N/A",
        "Offer URL": detail_url,
        "Images": " | ".join(img_list),
        "Page": page_index,
        "Item Index": item_index,
    }
    
    # توزيع الـ Tags بعد الدمج
    total_specs = max(15, len(tags))
    for n in range(1, total_specs + 1):
        row[f"specification{n}"] = tags[n - 1] if n <= len(tags) else ""

    return row


def fetch_page(
    session: requests.Session, base_payload: dict[str, Any], page_index: int
) -> tuple[list[dict[str, Any]], int, dict[str, Any], bool]:
    """
    Fetches a single page of offers using safe_request.
    Returns: (offers, page_count, raw_data, is_success)
    """
    payload = dict(base_payload)
    payload["pageIndex"] = page_index

    resp = safe_request(
        session,
        "POST",
        ENDPOINT,
        json_data=payload,
        headers=HEADERS,
        timeout=TIMEOUT_SECONDS,
        max_retries=MAX_RETRIES,
    )

    if resp is None:
        print(f"Network error on page {page_index}: No response received.")
        return [], 0, {}, False

    if resp.status_code == 403:
        print(f"HTTP 403 Forbidden on page {page_index}.")
        return [], 0, {}, False

    try:
        data = resp.json()
    except Exception:
        print(f"Page {page_index}: Non-JSON response (status {resp.status_code})")
        return [], 0, {}, False

    code = str(data.get("code", ""))
    if code in ("0000", "0", "200"):
        data_obj = data.get("data", {})
        if isinstance(data_obj, dict):
            offers = data_obj.get("list", []) or []
            page_count = int(data_obj.get("pageCount") or 1)
            return offers, page_count, data, True
    else:
        print(f"API Error on page {page_index}: code={code}, msg={data.get('msg')}")
        return [], 0, data, False

    return [], 0, {}, False


def main() -> None:
    payload = load_payload()

    print("=" * 60)
    print("IGV Marketplace API Scraper - Full Data & Profiles")
    print(f"Endpoint: {ENDPOINT}")
    print(f"Game: {GAME_NAME} | Category: {CATEGORY_NAME}")
    print(f"Tenant ID: {payload.get('tenantId')}")
    print(f"Mode: {'Pilot (max ' + str(MAX_PAGES) + ' pages)' if MAX_PAGES else 'Full scrape'}")
    print("=" * 60)

    session = requests.Session()

    # 1. جلب الصفحة الأولى
    print("\nFetching Page 1...")
    offers, page_count, raw_p1, p1_success = fetch_page(session, payload, 1)

    # حفظ استجابة الصفحة الأولى للفحص
    if raw_p1:
        DEBUG_RESPONSE_FILE.write_text(json.dumps(raw_p1, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Saved raw page 1 response to: {DEBUG_RESPONSE_FILE.resolve()}")

    if not p1_success or not offers:
        print("\nCould not retrieve offers from page 1. Inspect debug_response_page1.json.")
        return

    total_count = int(raw_p1.get("data", {}).get("totalCount") or 0)
    total_pages = min(page_count, MAX_PAGES) if MAX_PAGES else page_count
    print(f"API reports totalCount={total_count:,}; pageCount={page_count:,}. Scraping {total_pages:,} page(s)...")

    all_offers: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    failed_pages: list[int] = []
    empty_pages: list[int] = []
    failed_seller_count: int = 0
    failed_desc_count: int = 0
    empty_desc_count: int = 0

    for p_idx in range(1, total_pages + 1):
        if p_idx == 1:
            current_offers = offers
            p_success = p1_success
        else:
            print(f"\nFetching Page {p_idx}/{total_pages}...")
            current_offers, _, _, p_success = fetch_page(session, payload, p_idx)
            polite_sleep(REQUEST_DELAY_SECONDS)

        if not p_success:
            print(f"  [ERROR] Page {p_idx} failed to load.")
            failed_pages.append(p_idx)
            continue

        if len(current_offers) == 0:
            print(f"  [INFO] Page {p_idx} returned 0 offers.")
            empty_pages.append(p_idx)
            continue

        new_count = 0
        for i_idx, item in enumerate(current_offers, 1):
            if not isinstance(item, dict):
                continue

            # جلب بروفايل البائع مع تتبع الأخطاء
            seller_id = str(item.get("sellerId") or "")
            seller_info, seller_ok = fetch_seller_profile(session, seller_id) if seller_id else ({}, True)
            if not seller_ok:
                failed_seller_count += 1

            # جلب الوصف الكامل لكل عرض بدون HTML مع تتبع الأخطاء
            offer_id = str(item.get("id") or item.get("productSpuId") or "")
            target_desc_id = str(item.get("standardId") or offer_id or "")
            description, desc_status = fetch_offer_description(session, target_desc_id) if target_desc_id else ("N/A", "empty")
            if desc_status == "failed":
                failed_desc_count += 1
            elif desc_status == "empty":
                empty_desc_count += 1

            parsed = parse_offer(
                item,
                p_idx,
                i_idx,
                seller_info=seller_info,
                description=description
            )
            uid = parsed["Offer ID"] if parsed["Offer ID"] != "N/A" else f"{parsed['Title']}_{parsed['Seller ID']}"
            if uid in seen_ids:
                continue
            seen_ids.add(uid)
            all_offers.append(parsed)
            new_count += 1

        print(f"  -> Page {p_idx}: received {len(current_offers)} items, kept {new_count} unique offers. (Cached sellers: {len(seller_cache)} | Total HTTP Reqs: {REQUEST_STATS['total_requests']} | 429 Hits: {REQUEST_STATS['rate_limit_429_hits']})")

        # حفظ نسخة مؤقتة بعد كل صفحة بطريقة آمنة ضد القفل
        if all_offers:
            try:
                pd.DataFrame(all_offers).fillna("").to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
            except PermissionError:
                alt_csv = Path(f"igv_offers_checkpoint_{int(time.time())}.csv")
                try:
                    pd.DataFrame(all_offers).fillna("").to_csv(alt_csv, index=False, encoding="utf-8-sig")
                    print(f"  [Notice] {OUTPUT_CSV.name} is open/locked. Saved checkpoint to {alt_csv.name}")
                except Exception:
                    pass

    if not all_offers:
        print("\nNo offers were parsed.")
        return

    # تصدير ملف الإكسل النهائي
    df = pd.DataFrame(all_offers).fillna("")
    target_xlsx = OUTPUT_XLSX

    def save_excel(path: Path) -> None:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Offers", index=False)
            first_hit_str = (
                f"Req #{REQUEST_STATS['first_rate_limit_hit']['request_number']} at {REQUEST_STATS['first_rate_limit_hit']['time']}"
                if REQUEST_STATS["first_rate_limit_hit"]
                else "Never Triggered (Safe)"
            )
            summary_df = pd.DataFrame([
                ["Game Name", GAME_NAME],
                ["Category", CATEGORY_NAME],
                ["Total API Offers Available (totalCount)", total_count],
                ["Total API Pages Available (pageCount)", page_count],
                ["Target Pages To Fetch", total_pages],
                ["Successfully Fetched Pages", total_pages - len(failed_pages)],
                ["Failed Pages Count", len(failed_pages)],
                ["Failed Pages List", ", ".join(map(str, failed_pages)) if failed_pages else "None"],
                ["Empty Pages Count", len(empty_pages)],
                ["Total Unique Offers Exported", len(df)],
                ["Unique Sellers Profiled", len(seller_cache)],
                ["Seller Profile Requests Failed", failed_seller_count],
                ["Offer Description Requests Failed", failed_desc_count],
                ["Offers With No Description (Empty/None)", empty_desc_count],
                ["Total HTTP Requests Made", REQUEST_STATS["total_requests"]],
                ["HTTP 429 Rate Limit Hits", REQUEST_STATS["rate_limit_429_hits"]],
                ["First Rate Limit Encountered At", first_hit_str],
                ["Server Errors (5xx) Hits", REQUEST_STATS["server_error_5xx_hits"]],
            ], columns=["Metric", "Value"])
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            for sheet in writer.sheets.values():
                sheet.freeze_panes = "A2"
                for col in sheet.columns:
                    max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col[:100])
                    sheet.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 45)

    try:
        save_excel(target_xlsx)
    except PermissionError:
        target_xlsx = Path(f"igv_offers_pilot_{int(time.time())}.xlsx")
        save_excel(target_xlsx)
        print(f"[Notice] {OUTPUT_XLSX.name} was open/locked in Excel. Saved output to: {target_xlsx.name}")

    print("\n" + "=" * 60)
    print("SCRAPE & RATE LIMIT DIAGNOSTICS REPORT")
    print(f"Total Pages Attempted: {total_pages} | Offers Exported: {len(df):,}")
    print(f"Total HTTP Requests Dispatched: {REQUEST_STATS['total_requests']}")
    print(f"HTTP 429 Rate Limit Hits: {REQUEST_STATS['rate_limit_429_hits']}")
    if REQUEST_STATS['first_rate_limit_hit']:
        print(f"⚠️ First Rate Limit (429) hit at Request #{REQUEST_STATS['first_rate_limit_hit']['request_number']} (Time: {REQUEST_STATS['first_rate_limit_hit']['time']})")
    else:
        print("✅ No Rate Limit (HTTP 429) was triggered throughout the 50 pages!")
    if failed_pages:
        print(f"WARNING: {len(failed_pages)} page(s) failed during scraping: {failed_pages}")
    if failed_seller_count:
        print(f"WARNING: {failed_seller_count} seller profile request(s) failed.")
    if failed_desc_count:
        print(f"WARNING: {failed_desc_count} offer description request(s) failed.")
    print(f"Offers with no description provided by seller: {empty_desc_count}")
    print(f"DONE! Exported to: {target_xlsx.resolve()}")
    print(f"Checkpoint saved to: {OUTPUT_CSV.resolve()}")
    print(f"Unique sellers cached: {len(seller_cache)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
