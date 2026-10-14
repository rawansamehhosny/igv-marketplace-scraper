"""
IGV Categories Directory Generator (Without Offers)
---------------------------------------------------
Extracts all games and all of their available categories (Accounts, Items, Coins, Boosting, etc.)
from https://www.igv.com/all-games and saves a comprehensive directory Excel file.
Does NOT scrape individual offers — serves as the complete catalog/index of the platform.
"""

import sys
import re
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
import pandas as pd
from pathlib import Path

# Ensure UTF-8 output on Windows
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

OUTPUT_FILE = Path("igv_all_categories_directory.xlsx")

def build_directory():
    base_url = "https://www.igv.com"
    all_games_url = f"{base_url}/all-games"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Referer": "https://www.igv.com/",
    }

    print(f"جاري جلب دليل الألعاب والتصنيفات من: {all_games_url} ...")
    resp = requests.get(all_games_url, headers=headers, timeout=20)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, 'html.parser')
    info_cards = soup.find_all('div', class_=re.compile(r'\binfo\b'))
    print(f"تم العثور على {len(info_cards)} لعبة في الموقع. جاري تفكيك التصنيفات...")

    flat_categories = []
    games_summary = []
    cat_counts = {}

    for card in info_cards:
        brand_a = card.find('a', class_=re.compile(r'\bbrand\b'))
        h3 = card.find('h3')

        game_name = ""
        if h3 and h3.text.strip():
            game_name = h3.text.strip()
        elif brand_a and brand_a.text.strip():
            game_name = brand_a.text.strip()

        brand_href = brand_a.get('href', '').strip() if brand_a else ""
        game_slug = brand_href.strip('/')

        if not game_slug:
            continue
        if not game_name:
            game_name = game_slug.replace('-', ' ').title()

        game_main_url = urljoin(base_url, brand_href)
        cat_links = card.find_all('a', class_=re.compile(r'\bcategory\b'))

        categories_list = []
        has_accounts = False
        has_items = False
        accounts_url = "N/A"
        items_url = "N/A"

        for cat in cat_links:
            cat_name = cat.text.strip()
            cat_href = cat.get('href', '').strip()
            if not cat_name:
                continue

            categories_list.append(cat_name)
            cat_full_url = urljoin(base_url, cat_href)
            cat_lower = cat_name.lower()

            is_scraped_target = "No (Directory Only)"
            if "account" in cat_lower or "/accounts" in cat_href:
                has_accounts = True
                accounts_url = cat_full_url
                is_scraped_target = "Yes (Target - Accounts)"
            elif "item" in cat_lower or "/items" in cat_href:
                has_items = True
                items_url = cat_full_url
                is_scraped_target = "Yes (Target - Items)"

            cat_counts[cat_name] = cat_counts.get(cat_name, 0) + 1

            flat_categories.append({
                "Game Name": game_name,
                "Game Slug": game_slug,
                "Category Name": cat_name,
                "Category URL": cat_full_url,
                "Target For Scraping": is_scraped_target
            })

        games_summary.append({
            "Game Name": game_name,
            "Game Slug": game_slug,
            "Game Main URL": game_main_url,
            "Total Categories Count": len(categories_list),
            "Available Categories": ", ".join(categories_list) if categories_list else "None",
            "Has Accounts": "Yes" if has_accounts else "No",
            "Has Items": "Yes" if has_items else "No",
            "Accounts URL": accounts_url,
            "Items URL": items_url
        })

    # Summary by Category Name
    cat_summary = [
        {
            "Category Name": c_name,
            "Games Count": count,
            "Scraping Status": "Active (Offers Scraped)" if c_name.lower() in ("accounts", "items") else "Directory Index Only"
        }
        for c_name, count in sorted(cat_counts.items(), key=lambda x: x[1], reverse=True)
    ]

    print(f"\nإجمالي ألعاب الموقع: {len(games_summary)} لعبة")
    print(f"إجمالي روابط التصنيفات المستخرجة: {len(flat_categories)} رابط تصنيف")
    print(f"عدد الألعاب التي تملك Accounts: {sum(1 for g in games_summary if g['Has Accounts'] == 'Yes')}")
    print(f"عدد الألعاب التي تملك Items: {sum(1 for g in games_summary if g['Has Items'] == 'Yes')}")
    print(f"عدد الألعاب التي تملك الاثنين معاً: {sum(1 for g in games_summary if g['Has Accounts'] == 'Yes' and g['Has Items'] == 'Yes')}")

    # Write to Excel
    save_path = OUTPUT_FILE
    try:
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            pd.DataFrame(flat_categories).to_excel(writer, sheet_name="All Categories by Game", index=False)
            pd.DataFrame(games_summary).to_excel(writer, sheet_name="Games Overview", index=False)
            pd.DataFrame(cat_summary).to_excel(writer, sheet_name="Categories Summary", index=False)
    except PermissionError:
        save_path = Path("igv_all_categories_directory_new.xlsx")
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            pd.DataFrame(flat_categories).to_excel(writer, sheet_name="All Categories by Game", index=False)
            pd.DataFrame(games_summary).to_excel(writer, sheet_name="Games Overview", index=False)
            pd.DataFrame(cat_summary).to_excel(writer, sheet_name="Categories Summary", index=False)

    print(f"\n تم حفظ دليل التصنيفات الكامل بنجاح في: {save_path.resolve()}")
    return save_path

if __name__ == "__main__":
    build_directory()
