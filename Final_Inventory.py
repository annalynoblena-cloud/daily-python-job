import os
import pandas as pd
import re

def remove_illegal_chars(df):
    # Remove illegal control characters Excel can't handle
    illegal_pattern = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")
    return df.map(lambda x: illegal_pattern.sub("", str(x)) if isinstance(x, str) else x)


def process_latest_txt(SOURCE_FOLDER, DEST_FOLDER, DEST_FILE):
    # Get list of txt files with full paths
    files = [os.path.join(SOURCE_FOLDER, f) for f in os.listdir(SOURCE_FOLDER) if f.lower().endswith(".txt")]

    if not files:
        print(f"No TXT files found in source folder: {SOURCE_FOLDER}")
        return

    # Sort by creation time (latest first)
    files.sort(key=os.path.getctime, reverse=True)
    latest_file = files[0]

    print(f"\nProcessing latest file from: {SOURCE_FOLDER}")
    print(f"Latest file: {os.path.basename(latest_file)}")

    # Try different parsing methods
    try:
        df = pd.read_csv(latest_file, sep=None, engine="python", on_bad_lines="skip")
    except Exception as e1:
        print(f"Auto-detect failed: {e1}")
        try:
            df = pd.read_csv(latest_file, delim_whitespace=True, engine="python", on_bad_lines="skip")
        except Exception as e2:
            print(f"Whitespace split failed: {e2}")
            # Last resort: read raw lines into single column
            with open(latest_file, "r", encoding="utf-8", errors="ignore") as f:
                lines = f.readlines()
            df = pd.DataFrame({"RawText": [line.strip() for line in lines]})

    # 🧹 Clean illegal Excel characters
    df = remove_illegal_chars(df)

    # Ensure destination folder exists
    os.makedirs(DEST_FOLDER, exist_ok=True)

    # Save as Excel
    dest_path = os.path.join(DEST_FOLDER, DEST_FILE)
    df.to_excel(dest_path, index=False)

    print(f"✅ Converted '{os.path.basename(latest_file)}' → '{dest_path}'")


def main():
    # --- 1️⃣ INVENTORY Path ---
    process_latest_txt(
        SOURCE_FOLDER=r"\\sasan02\prd\SBU\INVENTORY",
        DEST_FOLDER=r"\\sasan02\prd\SBU\FINAL",
        DEST_FILE="FINAL_INVENTORY.xlsx"
    )

    # --- 2️⃣ OPENPO Path ---
    process_latest_txt(
        SOURCE_FOLDER=r"\\sasan02\prd\SBU\OPENPO",
        DEST_FOLDER=r"\\sasan02\prd\SBU\FINAL",
        DEST_FILE="FINAL_OPENPO.xlsx"
    )


if __name__ == "__main__":
    main()
