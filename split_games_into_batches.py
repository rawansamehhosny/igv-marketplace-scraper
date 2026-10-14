"""
Split IGV Games Directory into 50-game Batch Files
--------------------------------------------------
Reads `igv_games_list.xlsx` and divides the games into batches of 50,
saving individual files: `games_batch_1.xlsx`, `games_batch_2.xlsx`, etc.
"""

import sys
from pathlib import Path
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

SOURCE_FILE = Path("igv_games_list.xlsx")
BATCH_SIZE = 50
OUTPUT_DIR = Path("game_batches")

def split_games():
    if not SOURCE_FILE.exists():
        print(f"❌ Source file not found: {SOURCE_FILE}")
        return

    OUTPUT_DIR.mkdir(exist_ok=True)
    
    # Read Accounts sheet
    try:
        df = pd.read_excel(SOURCE_FILE, sheet_name="Accounts Games")
    except Exception:
        df = pd.read_excel(SOURCE_FILE, sheet_name=0)

    total_games = len(df)
    total_batches = (total_games + BATCH_SIZE - 1) // BATCH_SIZE

    print(f"🎮 Total Games Found: {total_games}")
    print(f"📦 Batch Size: {BATCH_SIZE} games per file")
    print(f"📁 Output Directory: {OUTPUT_DIR.resolve()}\n")

    for b_idx in range(total_batches):
        start = b_idx * BATCH_SIZE
        end = min(start + BATCH_SIZE, total_games)
        df_batch = df.iloc[start:end]
        
        batch_filename = OUTPUT_DIR / f"games_batch_{b_idx + 1}_{start + 1}_to_{end}.xlsx"
        with pd.ExcelWriter(batch_filename, engine="openpyxl") as writer:
            df_batch.to_excel(writer, sheet_name="Accounts Games", index=False)
            
        print(f"  ✅ Batch {b_idx + 1}/{total_batches}: {len(df_batch)} games saved -> {batch_filename.name}")

    print("\n✨ All batch files created successfully!")

if __name__ == "__main__":
    split_games()
