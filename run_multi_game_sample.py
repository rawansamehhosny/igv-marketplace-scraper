"""
IGV Multi-Category Scraper (Accounts & Items & Multi-Game)
----------------------------------------------------------
Supports both standard and marketplace categories seamlessly:
- Accounts: uses `ACCOUNTS_ENDPOINT` (/product-market/product/page)
- Items / Coins / Other: uses `ITEMS_ENDPOINT` (/v2/product/standard) with mixPageResp parser
"""

from __future__ import annotations

import sys
import os
import time
import random
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

# Ensure UTF-8 console output on Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ---------- CONFIGURATION ----------
ACCOUNTS_ENDPOINT = "https://www.igv.com/api/product-search/product-market/product/page"
ITEMS_ENDPOINT = "https://www.igv.com/api/product-search/v2/product/standard"
SELLER_STORE_ENDPOINT = "https://www.igv.com/api/operation/seller/store"
PRODUCT_DETAIL_ENDPOINT = "https://www.igv.com/api/product-search/v2/product/standard/"
SITE_BASE_URL = "https://www.igv.com/"

TENANT_ID = "3332001"
SITE_ID = "1628721356376162310"
USD_CURRENCY_ID = "1598942409803481063"

# ---------- BATCHING & SAFETY CONTROLS ----------
PAGES_PER_GAME = 1
TIMEOUT_SECONDS = 25
MAX_RETRIES = 3
REQUEST_DELAY_SECONDS = 1.5   # Polite delay between page requests (was 0.8)
ITEM_REQUEST_DELAY = 0.5      # Polite delay between item-level requests to prevent 429/403 (was 0.35)

# نظام تقسيم العمل (Batching)
USE_EXCEL_TARGETS = True      # True لقراءة قائمة الألعاب من ملف إكسل، False لاستخدام عينة test_targets
SOURCE_GAMES_FILE = "igv_games_list.xlsx"
BATCH_SIZE = 50               # عدد الألعاب في الدفعة الواحدة (الافتراضي: 50)
BATCH_NUMBER = 1              # رقم الدفعة: 1 = أول 50 لعبة، 2 = من 51 إلى 100، وهكذا.. (0 = سحب كل الألعاب)
AUTO_RESUME = True            # تخطي أي لعبة تم سحبها مسبقاً وتوجد في ملف الـ Checkpoint

COMMON_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "lang": "en",
    "timezone": "Africa/Cairo",
    "deviceid": "51a2d977eca7ac4e6802995744a18812",
    "ab-test-group": "1",
    "tenant-id": TENANT_ID,
    "origin": "https://www.igv.com",
}

# Cache for seller store profiles to eliminate duplicate HTTP calls
seller_cache: dict[str, dict[str, Any]] = {}

def polite_sleep(base_seconds: float, jitter_ratio: float = 0.25) -> None:
    """Sleeps with slight random jitter around base_seconds to avoid bot detection patterns."""
    low = max(0.1, base_seconds * (1.0 - jitter_ratio))
    high = base_seconds * (1.0 + jitter_ratio)
    time.sleep(random.uniform(low, high))

# ---------- HELPER FUNCTIONS ----------

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
    for attempt in range(max_retries + 1):
        try:
            resp = session.request(
                method,
                url,
                params=params,
                json=json_data,
                headers=headers or COMMON_HEADERS,
                timeout=timeout,
            )
            if resp.status_code == 429:
                retry_after_str = resp.headers.get("Retry-After")
                wait_sec = float(retry_after_str) if retry_after_str else float(2 ** attempt + 2)
                time.sleep(wait_sec)
                continue
            if resp.status_code == 403:
                print(f"\n  🛑 [Cloudflare / WAF 403 Forbidden] Blocked request on: {url}")
                if attempt < max_retries:
                    wait_sec = float(5.0 * (attempt + 1))
                    print(f"  🛑 Waiting {wait_sec:.1f}s cooldown before retrying...")
                    time.sleep(wait_sec)
                    continue
                return resp
            if 500 <= resp.status_code < 600:
                time.sleep(float(2 ** attempt))
                continue
            return resp
        except requests.RequestException:
            if attempt >= max_retries:
                return None
            time.sleep(float(2 ** attempt))
    return None


def resolve_game_meta(session: requests.Session, game_url: str) -> tuple[str | None, str | None]:
    html_headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    resp = safe_request(session, "GET", game_url, headers=html_headers)
    if resp is None or resp.status_code != 200:
        return None, None
    
    # 1. Search in Schema.org ld+json (fastest and most accurate)
    m_sku = re.search(r'"sku"\s*:\s*"([0-9]+)&([0-9]+)"', resp.text)
    if m_sku:
        return m_sku.group(1), m_sku.group(2)
        
    # 2. Search in Nuxt state
    m_brand = re.search(r'brandId["\':\s]+([0-9]+)', resp.text)
    m_cat = re.search(r'tenantCategoryId3?["\':\s]+([0-9]+)', resp.text)
    if m_brand and m_cat:
        return m_brand.group(1), m_cat.group(1)
        
    return None, None


def clean_html_text(raw_html: str) -> str:
    if not raw_html or not str(raw_html).strip():
        return "N/A"
    clean = re.sub(r'</?(?:p|div|br|li|h[1-6]|tr)[^>]*>', '\n', str(raw_html), flags=re.IGNORECASE)
    clean = re.sub(r'<[^>]+>', ' ', clean)
    clean = re.sub(r'[ \t]+', ' ', clean)
    clean = re.sub(r'\n\s*\n+', '\n', clean).strip()
    return clean if clean else "N/A"


def calculate_seller_tenure(register_time_str: str) -> str:
    if not register_time_str or str(register_time_str).strip() in ("", "N/A", "None"):
        return "N/A"
    try:
        dt = datetime.strptime(str(register_time_str).split()[0], "%Y-%m-%d")
        days = (datetime.now() - dt).days
        if days < 0:
            return "Recently joined"
        years = days // 365
        months = (days % 365) // 30
        if years >= 1:
            return f"{years} year(s) {months} month(s)" if months > 0 else f"{years} year(s)"
        elif months >= 1:
            return f"{months} month(s)"
        else:
            return f"{days} day(s)"
    except Exception:
        return "N/A"


def fetch_seller_profile(session: requests.Session, seller_id: str) -> tuple[dict[str, Any], bool]:
    if not seller_id or str(seller_id).strip() in ("", "None", "N/A"):
        return {}, True
    seller_id_str = str(seller_id).strip()
    if seller_id_str in seller_cache:
        return seller_cache[seller_id_str], True

    polite_sleep(ITEM_REQUEST_DELAY)
    headers = dict(COMMON_HEADERS)
    headers["referer"] = f"https://www.igv.com/store/{seller_id_str}"
    params = {
        "sellerId": seller_id_str,
        "tenantId": TENANT_ID,
        "siteId": SITE_ID,
    }
    resp = safe_request(session, "GET", SELLER_STORE_ENDPOINT, params=params, headers=headers)
    if resp is None or resp.status_code != 200:
        return {}, False
    try:
        data = resp.json()
        if str(data.get("code")) in ("0000", "0", "200") and isinstance(data.get("data"), dict):
            seller_cache[seller_id_str] = data["data"]
            return data["data"], True
    except Exception:
        pass
    return {}, False


def fetch_offer_description(session: requests.Session, offer_id: str, game_url: str) -> tuple[str, str]:
    if not offer_id or str(offer_id).strip() in ("", "None", "N/A"):
        return "N/A", "empty"
    polite_sleep(ITEM_REQUEST_DELAY)
    url = f"{PRODUCT_DETAIL_ENDPOINT}{str(offer_id).strip()}"
    headers = dict(COMMON_HEADERS)
    headers["referer"] = f"{game_url.rstrip('/')}/{offer_id}"
    payload = {"siteId": SITE_ID, "tenantId": TENANT_ID}
    
    resp = safe_request(session, "POST", url, json_data=payload, headers=headers)
    if resp is None or resp.status_code != 200:
        return "N/A", "failed"
    try:
        data = resp.json()
        if str(data.get("code")) in ("0000", "0", "200"):
            raw_desc = (data.get("data") or {}).get("noResp", {}).get("description")
            if raw_desc and str(raw_desc).strip():
                cleaned = clean_html_text(str(raw_desc))
                return (cleaned, "ok") if cleaned != "N/A" else ("N/A", "empty")
            return "N/A", "empty"
    except Exception:
        pass
    return "N/A", "failed"


def extract_tags(item: dict[str, Any]) -> list[str]:
    attrs = item.get("productSpuCategoryAttrs") or []
    grouped_tags: dict[str, list[str]] = {}
    for attr in attrs:
        if isinstance(attr, dict):
            name = str(attr.get("categoryAttrName") or "").strip()
            val = str(attr.get("categoryAttrValueName") or "").strip()
            if name and val:
                grouped_tags.setdefault(name, []).append(val)
            elif val and not name:
                grouped_tags.setdefault("Other", []).append(val)
    tags: list[str] = []
    for name, vals in grouped_tags.items():
        if name == "Other":
            tags.extend(vals)
        else:
            tags.append(f"{name}//{'--'.join(vals)}")
    return tags


def extract_images(item: dict[str, Any]) -> list[str]:
    img_list: list[str] = []
    # Both standard items and account items supported
    imgs = item.get("productStandardImages") or item.get("productSpuImages") or []
    for img in imgs:
        if isinstance(img, dict) and img.get("imageUrl"):
            url = str(img["imageUrl"]).strip()
            if url and url not in img_list:
                img_list.append(url)
    if not img_list and item.get("imageUrl"):
        img_list.append(str(item["imageUrl"]).strip())
    return img_list


def parse_offer(
    item: dict[str, Any],
    game_name: str,
    category_name: str,
    game_url: str,
    page_index: int,
    item_index: int,
    seller_info: dict[str, Any],
    description: str,
) -> dict[str, Any]:
    offer_id = str(item.get("id") or item.get("productSpuId") or "")
    detail_url = urljoin(SITE_BASE_URL, f"product/{offer_id}") if offer_id else game_url
    tags = extract_tags(item)
    img_list = extract_images(item)

    # Offer Type
    raw_type = str(item.get("productType") or item.get("type") or "").strip()
    if raw_type in ("1", "sale", "Sale"):
        offer_type = "Sale"
    elif raw_type in ("2", "rent", "Rent", "rental"):
        offer_type = "Rent"
    else:
        offer_type = "Sale"

    # Delivery speed
    delivery_code = str(item.get("deliveryType") or "").strip()
    if delivery_code == "1":
        delivery_speed = "Instant"
    elif delivery_code == "2":
        delivery_speed = "10 Min"
    elif delivery_code == "3":
        delivery_speed = "Manual / Specified"
    else:
        delivery_speed = delivery_code if delivery_code not in ("None", "") else "N/A"

    # Warranty / Guarantee
    hours = item.get("afterSalesHours") or item.get("categoryGuaranteeHours")
    days = item.get("categoryGuaranteeDays")
    if hours and str(hours).strip() not in ("None", "0"):
        warranty = f"{hours}h Warranty"
    elif days and str(days).strip() not in ("None", "0"):
        warranty = f"{days}d Warranty"
    elif item.get("isGuaranteedCompensation") is True:
        warranty = "Guaranteed"
    else:
        warranty = "N/A"

    # Seller Name: prefer storeName from store profile
    seller_name = seller_info.get("storeName") or seller_info.get("sellerName") or item.get("sellerStoreName") or item.get("sellerName") or "N/A"

    # Seller level & tier
    seller_level = seller_info.get("sellerStarName") or item.get("sellerStarName") or "N/A"
    seller_tier = seller_info.get("tierName") or item.get("tierName") or "N/A"

    # Seller rating & reviews
    raw_rate = seller_info.get("reviewRate")
    review_rate_str = "N/A"
    if raw_rate is not None and str(raw_rate).strip() not in ("", "N/A", "None"):
        try:
            val = float(raw_rate)
            review_rate_str = f"{val * 100:.1f}%" if val <= 1.0 else f"{val:.1f}%"
        except (ValueError, TypeError):
            review_rate_str = str(raw_rate)

    review_count = str(seller_info.get("reviewCount")) if seller_info.get("reviewCount") is not None else "N/A"
    nice_review_count = str(seller_info.get("niceReviewCount")) if seller_info.get("niceReviewCount") is not None else "N/A"
    avg_score = str(seller_info.get("averageScore")) if seller_info.get("averageScore") is not None else "N/A"
    order_count = str(seller_info.get("orderCount")) if seller_info.get("orderCount") is not None else "N/A"
    sales_vol = str(seller_info.get("salesVolume")) if seller_info.get("salesVolume") is not None else "N/A"
    delivery_score = str(seller_info.get("deliverySpeedScore")) if seller_info.get("deliverySpeedScore") is not None else "N/A"

    is_kyc = seller_info.get("isCompletedEKyc")
    verified_str = "Yes" if is_kyc is True else ("No" if is_kyc is False else "N/A")
    reg_time = str(seller_info.get("registerTime") or "N/A")
    seller_tenure = calculate_seller_tenure(reg_time)

    # Price in USD
    skus = item.get("productSkus") or []
    price = "N/A"
    for sku in skus:
        if isinstance(sku, dict) and (str(sku.get("currencyCode")).upper() == "USD" or str(sku.get("currencyId")) == USD_CURRENCY_ID):
            price = sku.get("price")
            break
    if price in ("N/A", None):
        price = item.get("price") or item.get("salePrice") or "N/A"
    if price not in ("N/A", None, ""):
        try:
            val = float(price)
            price = f"{val:.2f}".rstrip("0").rstrip(".") if "." in f"{val:.2f}" else f"{val:.0f}"
        except (ValueError, TypeError):
            pass

    row: dict[str, Any] = {
        "Game Name": game_name,
        "Category": category_name,
        "Offer Type": offer_type,
        "Title": str(item.get("title") or item.get("productSpuName") or "N/A"),
        "Price": str(price),
        "Currency": "USD",
        "Delivery Speed": str(delivery_speed),
        "Warranty": str(warranty),
        "Seller ID": str(item.get("sellerId") or "N/A"),
        "Seller Name": str(seller_name),
        "Seller Level": str(seller_level),
        "Seller Tier": str(seller_tier),
        "Seller Tenure": str(seller_tenure),
        "Seller Member Since": str(reg_time),
        "Seller Review Count": str(review_count),
        "Seller Positive Reviews": str(nice_review_count),
        "Seller Avg Rating": str(avg_score),
        "Seller Satisfaction Rate": str(review_rate_str),
        "Seller Completed Orders": str(order_count),
        "Seller Sales Volume": str(sales_vol),
        "Seller Verified (KYC)": str(verified_str),
        "Seller Delivery Speed": str(delivery_score),
        "Publish Date": str(item.get("publishDate") or item.get("createTime") or "N/A"),
        "Description": str(description if description and description != "N/A" else "N/A"),
        "Offer ID": str(offer_id or "N/A"),
        "Offer URL": str(detail_url),
        "Images": " | ".join(img_list) if img_list else "N/A",
        "Page": page_index,
        "Item Index": item_index,
    }

    # Dynamic specifications
    total_specs = max(15, len(tags))
    for n in range(1, total_specs + 1):
        row[f"specification{n}"] = tags[n - 1] if n <= len(tags) else ""

    return row


def scrape_game(
    session: requests.Session,
    game_name: str,
    category_name: str,
    game_url: str,
    max_pages: int = PAGES_PER_GAME,
    max_items_per_page: int | None = None,
) -> list[dict[str, Any]]:
    print(f"\n{'='*60}")
    print(f"🎮 [GAME] {game_name} | 🏷️ [CATEGORY] {category_name}")
    print(f"🔗 [URL] {game_url}")
    
    brand_id, cat_id = resolve_game_meta(session, game_url)
    if not brand_id or not cat_id:
        print(f"  ❌ Could not resolve brandId/categoryId for {game_name}.")
        return []

    print(f"  ✅ Resolved -> brandId: {brand_id} | tenantCategoryId3: {cat_id}")

    headers = dict(COMMON_HEADERS)
    headers["referer"] = game_url
    
    # تحديد هل الكاتيجوري ده Accounts ولا حاجة تانية (زي Items)
    is_standard_api = category_name.lower() != "accounts"
    current_endpoint = ITEMS_ENDPOINT if is_standard_api else ACCOUNTS_ENDPOINT

    # الـ Payload الأساسي المشترك
    payload = {
        "tenantId": TENANT_ID,
        "isTenantOnSale": True,
        "pageIndex": 1,
        "pageSize": 18 if is_standard_api else 16,
        "brandIds": [str(brand_id)],
        "isShowAttrTag": False,
        "isShowAttr": True,
        "isShowGuarantee": True,
        "productCategoryAttrs": [],
        "keywords": "",
        "tenantCategoryId3": str(cat_id),
        "automationKeywords": "",
        "sorts": [{"field": "bestMatchScore", "isAsc": False}],
        "isShowSellerInfo": False if is_standard_api else True,
        "isShowSellerStar": False if is_standard_api else True,
        "siteId": SITE_ID,
        "toCurrencyId": USD_CURRENCY_ID,
    }

    # إضافة الخصائص الخاصة بالـ Items
    if is_standard_api:
        payload.update({
            "isGetSellerInfo": True,
            "isGetGuaranteeInfo": False,
            "isGetDeliveryDetail": False,
            "isGetMixCategoryAttr": True,
            "isGetSellerAuthLevelOperationRecord": False,
        })

    offers_collected: list[dict[str, Any]] = []

    for page_idx in range(1, max_pages + 1):
        payload["pageIndex"] = page_idx
        print(f"  📄 Fetching Page {page_idx}/{max_pages}...")
        
        resp = safe_request(session, "POST", current_endpoint, json_data=payload, headers=headers)
        if resp is None or resp.status_code != 200:
            print(f"    ❌ Failed to load page {page_idx}.")
            continue

        data = resp.json()
        if str(data.get("code")) not in ("0000", "0", "200"):
            print(f"    ❌ API Error: {data.get('msg')}")
            break

        data_obj = data.get("data") or {}
        
        # بنسحب الداتا من المكان الصح حسب نوع الـ API
        if is_standard_api:
            mix_resp = data_obj.get("mixPageResp") or {}
            items = mix_resp.get("list") or []
            total_items = mix_resp.get("totalCount") or 0
        else:
            items = data_obj.get("list") or []
            total_items = data_obj.get("totalCount") or 0
        
        if not items:
            print(f"    ℹ️ Page {page_idx} returned 0 offers. End of results.")
            break

        if max_items_per_page and len(items) > max_items_per_page:
            items = items[:max_items_per_page]

        print(f"    📦 Page {page_idx} contains {len(items)} offers (Game total available: {total_items:,}). Enriching...")
        
        for item_idx, item in enumerate(items, 1):
            seller_id = str(item.get("sellerId") or "")
            seller_info, _ = fetch_seller_profile(session, seller_id) if seller_id else ({}, True)

            offer_id = str(item.get("id") or item.get("productSpuId") or "")
            target_desc_id = str(item.get("standardId") or offer_id or "")
            desc, _ = fetch_offer_description(session, target_desc_id, game_url) if target_desc_id else ("N/A", "empty")

            parsed = parse_offer(
                item=item,
                game_name=game_name,
                category_name=category_name,
                game_url=game_url,
                page_index=page_idx,
                item_index=item_idx,
                seller_info=seller_info,
                description=desc,
            )
            offers_collected.append(parsed)

        polite_sleep(REQUEST_DELAY_SECONDS)

    print(f"  ✨ Done for {game_name} ({category_name}): {len(offers_collected)} offers scraped.")
    return offers_collected


def main() -> None:
    print("=" * 70)
    print("🚀 IGV Multi-Game & Multi-Category Scraper (Accounts & Items Sample)")
    print("=" * 70)

    session = requests.Session()
    all_results: list[dict[str, Any]] = []

    # Selected multi-game sample: Fortnite, Clash of Clan, GTA V, Genshin Impact
    test_targets = [
        {"Game": "Fortnite", "Category": "Accounts", "URL": "https://www.igv.com/fortnite/accounts"},
        {"Game": "Fortnite", "Category": "Items", "URL": "https://www.igv.com/fortnite/items"},
        {"Game": "Clash of Clan", "Category": "Accounts", "URL": "https://www.igv.com/clash-of-clan/accounts"},
        {"Game": "Clash of Clan", "Category": "Items", "URL": "https://www.igv.com/clash-of-clan/items"},
        {"Game": "GTA V", "Category": "Accounts", "URL": "https://www.igv.com/gta-v/accounts"},
        {"Game": "GTA V", "Category": "Items", "URL": "https://www.igv.com/gta-v/items"},
        {"Game": "Genshin Impact", "Category": "Accounts", "URL": "https://www.igv.com/genshin-impact/accounts"},
        {"Game": "Genshin Impact", "Category": "Items", "URL": "https://www.igv.com/genshin-impact/items"},
    ]

    # تحميل الأهداف حسب الـ Batching والملف المصدر
    targets: list[dict[str, str]] = []
    source_file = Path(SOURCE_GAMES_FILE)
    if USE_EXCEL_TARGETS and source_file.exists():
        try:
            df_source = pd.read_excel(source_file, sheet_name=0)
            for _, row in df_source.iterrows():
                g_name = str(row.get("Game Name") or "").strip()
                g_url = str(row.get("Game URL") or "").strip()
                if g_name and g_url and g_url.startswith("http"):
                    # تحديد التصنيف بناء على الرابط أو الافتراضي Accounts
                    cat_name = "Accounts" if "/accounts" in g_url.lower() else ("Items" if "/items" in g_url.lower() else "Accounts")
                    targets.append({"Game": g_name, "Category": cat_name, "URL": g_url})
            print(f"📋 Loaded {len(targets)} game targets from {source_file.name}")
        except Exception as e:
            print(f"⚠️ Warning loading {source_file}: {e}. Falling back to default sample targets.")
            targets = test_targets
    else:
        targets = test_targets

    # تطبيق التجزئة (Batching)
    if BATCH_NUMBER > 0 and len(targets) > BATCH_SIZE:
        start_idx = (BATCH_NUMBER - 1) * BATCH_SIZE
        end_idx = min(start_idx + BATCH_SIZE, len(targets))
        selected_targets = targets[start_idx:end_idx]
        print(f"📦 Running Batch #{BATCH_NUMBER} (Games {start_idx + 1} to {end_idx} of {len(targets)} total)")
        checkpoint_csv = Path(f"igv_offers_batch_{BATCH_NUMBER}_checkpoint.csv")
        excel_path = Path(f"igv_offers_batch_{BATCH_NUMBER}.xlsx")
    else:
        selected_targets = targets
        checkpoint_csv = Path("multi_game_accounts_and_items_checkpoint.csv")
        excel_path = Path("multi_game_accounts_and_items_sample.xlsx")

    # نظام الاستئناف التلقائي (Auto-Resume)
    completed_keys: set[tuple[str, str]] = set()
    if AUTO_RESUME and checkpoint_csv.exists():
        try:
            df_cp = pd.read_csv(checkpoint_csv, encoding="utf-8-sig")
            if not df_cp.empty and "Game Name" in df_cp.columns and "Category" in df_cp.columns:
                all_results = df_cp.to_dict(orient="records")
                completed_keys = set(zip(df_cp["Game Name"].astype(str), df_cp["Category"].astype(str)))
                print(f"🔄 [AUTO-RESUME] Found {len(all_results)} existing offers ({len(completed_keys)} targets) in {checkpoint_csv.name}. Will skip them.")
        except Exception as e:
            print(f"⚠️ [AUTO-RESUME] Could not parse existing checkpoint: {e}")

    start_time = time.time()

    for idx, data in enumerate(selected_targets, 1):
        target_key = (data["Game"], data["Category"])
        if AUTO_RESUME and target_key in completed_keys:
            print(f"⏩ [SKIP / AUTO-RESUME] ({idx}/{len(selected_targets)}) {data['Game']} ({data['Category']}) already scraped. Skipping...")
            continue

        offers = scrape_game(
            session=session,
            game_name=data["Game"],
            category_name=data["Category"],
            game_url=data["URL"],
            max_pages=PAGES_PER_GAME
        )
        all_results.extend(offers)
        completed_keys.add(target_key)
        
        # Save incremental CSV checkpoint
        if all_results:
            df_temp = pd.DataFrame(all_results)
            df_temp.to_csv(checkpoint_csv, index=False, encoding="utf-8-sig")

    if not all_results:
        print("\n❌ No offers were collected.")
        return

    elapsed = time.time() - start_time
    df = pd.DataFrame(all_results)
    
    # Fill empty specs with ""
    spec_cols = [c for c in df.columns if c.startswith("specification")]
    for sc in spec_cols:
        df[sc] = df[sc].fillna("")

    save_excel_path = excel_path
    try:
        with pd.ExcelWriter(save_excel_path, engine="openpyxl") as writer:
            # 1. Sheet for all offers
            df.to_excel(writer, sheet_name="All Combined Offers", index=False)
            # 2. Sheet for Accounts
            df[df["Category"] == "Accounts"].to_excel(writer, sheet_name="Accounts Offers", index=False)
            # 3. Sheet for Items
            df[df["Category"] == "Items"].to_excel(writer, sheet_name="Items Offers", index=False)
            # 4. Sheets for each game
            for g_name in df["Game Name"].unique():
                sheet_safe_name = re.sub(r'[\\/*?:\[\]]', '', str(g_name))[:30]
                df[df["Game Name"] == g_name].to_excel(writer, sheet_name=sheet_safe_name, index=False)
    except PermissionError:
        save_excel_path = Path("multi_game_accounts_and_items_sample_new.xlsx")
        with pd.ExcelWriter(save_excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="All Combined Offers", index=False)
            df[df["Category"] == "Accounts"].to_excel(writer, sheet_name="Accounts Offers", index=False)
            df[df["Category"] == "Items"].to_excel(writer, sheet_name="Items Offers", index=False)
            for g_name in df["Game Name"].unique():
                sheet_safe_name = re.sub(r'[\\/*?:\[\]]', '', str(g_name))[:30]
                df[df["Game Name"] == g_name].to_excel(writer, sheet_name=sheet_safe_name, index=False)

    print("\n" + "=" * 70)
    print("🎉 MULTI-GAME & MULTI-CATEGORY SCRAPING TEST COMPLETED!")
    print("=" * 70)
    print(f"Total Offers Scraped: {len(all_results)} offers across 4 games (Accounts + Items) in {elapsed:.1f}s")
    print(f"Excel File: {save_excel_path.resolve()}")
    print(f"CSV Checkpoint: {checkpoint_csv.resolve()}")
    print("\nBreakdown by Game & Category:")
    for (g_name, cat), grp in df.groupby(["Game Name", "Category"]):
        print(f"  • {g_name:<16} | {cat:<10}: {len(grp):>2} offers")
    print("=" * 70)


if __name__ == "__main__":
    main()
