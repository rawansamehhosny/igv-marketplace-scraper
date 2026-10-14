import sys
import re
import json
import time
import random
from pathlib import Path
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup
import pandas as pd

# ---------- CONFIGURATION ----------
REQUEST_DELAY_SECONDS = 1.5  # Polite delay between page requests
ITEM_REQUEST_DELAY = 0.5   # Polite delay for item-level operations (unused here)

def polite_sleep(base_seconds: float, jitter_ratio: float = 0.25) -> None:
    """Sleep with random jitter around base_seconds to mimic human behavior."""
    low = max(0.1, base_seconds * (1.0 - jitter_ratio))
    high = base_seconds * (1.0 + jitter_ratio)
    time.sleep(random.uniform(low, high))

# Fix Windows console UTF-8 display
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

def get_all_games():
    """
    استخراج الفهرس الكامل لجميع ألعاب قسم الحسابات (Accounts) في موقع IGV (579 لعبة).
    يقوم بالسحب المباشر من صفحة التصنيف الرسمية: https://www.igv.com/categories/accounts
    ويستخرج: اسم اللعبة، عدد العروض المتاحة، الرابط، ومعرف اللعبة.
    """
    base_url = "https://www.igv.com"
    accounts_url = f"{base_url}/categories/accounts"
    local_cache_file = Path("categories_accounts_raw.html")

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.igv.com/",
    }

    html_content = ""

    # 1. محاولة استخدام النسخة المحلية المخبأة أولاً لو وجدت لتوفير الوقت والشبكة
    if local_cache_file.exists() and local_cache_file.stat().st_size > 100000:
        print(f"📦 تم العثور على نسخة HTML مخبأة ({local_cache_file.name}). جاري قراءتها...")
        html_content = local_cache_file.read_text(encoding="utf-8")
    else:
        print(f"🌐 جاري طلب صفحة الألعاب من: {accounts_url} ...")
        session = requests.Session()
        for attempt in range(1, 4):
            try:
                polite_sleep(REQUEST_DELAY_SECONDS)
                resp = session.get(accounts_url, headers=headers, timeout=45)
                if resp.status_code == 200 and len(resp.text) > 50000:
                    html_content = resp.text
                    local_cache_file.write_text(html_content, encoding="utf-8")
                    print(f"✅ تم تحميل الصفحة بنجاح وحفظ نسخة محلية.")
                    break
            except Exception as e:
                print(f"  ⚠️ المحاولة {attempt} فشلت: {e}. إعادة المحاولة بعد ثانيتين...")
                time.sleep(2)
            except Exception as e:
                print(f"  ⚠️ المحاولة {attempt} فشلت: {e}. إعادة المحاولة بعد ثانيتين...")
                time.sleep(2)

    if not html_content:
        print("❌ تعذر تحميل الصفحة. يرجى التأكد من اتصال الإنترنت.")
        return

    print("🔍 جاري تحليل وتفكيك بيانات كروت الألعاب من الـ HTML...")
    soup = BeautifulSoup(html_content, "html.parser")
    pattern = re.compile(r"^/([^/]+)/accounts/?$")

    games_list = []
    seen_slugs = set()

    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()
        m = pattern.match(href)
        if not m:
            continue

        slug = m.group(1)
        if slug in ("categories", "all-games", "user", "help", "store") or slug in seen_slugs:
            continue
        seen_slugs.add(slug)

        # 1. اسم اللعبة النظيف (من عنصر brand-name الداخلي أو النص)
        name_el = a.find(class_=re.compile(r"brand-name")) or (a.parent.find(class_=re.compile(r"brand-name")) if a.parent else None)
        if name_el and name_el.get_text(strip=True):
            game_name = name_el.get_text(strip=True)
        else:
            raw_text = a.get_text(separator=" ").strip()
            game_name = re.sub(r"\s*\d+\s*Offers?$", "", raw_text, flags=re.IGNORECASE).strip()
            if not game_name:
                game_name = slug.replace("-", " ").title()

        # 2. عدد العروض المتوفرة (Available Offers)
        offer_el = a.find(class_=re.compile(r"brand-offer")) or (a.parent.find(class_=re.compile(r"brand-offer")) if a.parent else None)
        offers_count = 0
        if offer_el:
            m_num = re.search(r"(\d+)", offer_el.get_text(strip=True))
            if m_num:
                offers_count = int(m_num.group(1))
        else:
            m_num = re.search(r"(\d+)\s*Offers?", a.get_text())
            if m_num:
                offers_count = int(m_num.group(1))

        # 3. أيقونة اللعبة (Icon URL)
        img = a.find("img") or (a.parent.find("img") if a.parent else None)
        icon_url = img.get("src") or img.get("data-src") or "N/A" if img else "N/A"

        full_url = urljoin(base_url, href)

        games_list.append({
            "Game Name": game_name,
            "Game Slug": slug,
            "Available Offers": offers_count,
            "Game URL": full_url,
            "Accounts URL": full_url,
            "All Categories": "Accounts",
            "Icon URL": icon_url,
        })

    print(f"\n🎉 إجمالي الألعاب المستخرجة بنجاح: {len(games_list)} لعبة!")

    if not games_list:
        print("❌ لم يتم العثور على ألعاب.")
        return

    # حفظ في الملفات
    df = pd.DataFrame(games_list)

    # 1. تحديث الملف الرئيسي المستهدف igv_games_list.xlsx
    main_excel = Path("igv_games_list.xlsx")
    with pd.ExcelWriter(main_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Accounts Games", index=False)
    print(f"✅ تم تحديث الملف الأساسي: {main_excel.resolve()} ({len(df)} لعبة)")

    # 2. حفظ نسخة بالاسم igv_579_games.xlsx
    backup_excel = Path("igv_579_games.xlsx")
    df.to_excel(backup_excel, index=False)
    print(f"✅ تم حفظ نسخة إضافية: {backup_excel.resolve()}")

    # إحصائيات سريعة لأهم الألعاب
    print("\nأكثر 5 ألعاب توفراً للعروض في IGV:")
    top_5 = df.sort_values(by="Available Offers", ascending=False).head(5)
    for _, row in top_5.iterrows():
        print(f"  • {row['Game Name']:<20}: {row['Available Offers']:,} أوفر")

if __name__ == "__main__":
    get_all_games()
