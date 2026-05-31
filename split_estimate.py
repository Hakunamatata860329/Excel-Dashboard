import pandas as pd
from pathlib import Path

INPUT_DIR  = Path(__file__).parent / "data" / "input"
OUTPUT_DIR = Path(__file__).parent / "data" / "output"

def run():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.read_excel(INPUT_DIR / "ESTIMATE.xlsx", sheet_name="in")
    is_with = (
        df["Estimate"].notna() &
        df["TimeSpend"].notna() &
        (df["Type"] == "Task")
    )
    df[is_with].to_excel(OUTPUT_DIR / "ESTIMATE_with.xlsx", index=False)
    df[~is_with].to_excel(OUTPUT_DIR / "ESTIMATE_without.xlsx", index=False)
    print(f"ESTIMATE_with.xlsx:    {is_with.sum()} rows")
    print(f"ESTIMATE_without.xlsx: {(~is_with).sum()} rows")

if __name__ == "__main__":
    run()
