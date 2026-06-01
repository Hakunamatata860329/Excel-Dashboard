import pandas as pd
from pathlib import Path

INPUT_DIR  = Path(__file__).parent / "data" / "input"
OUTPUT_DIR = Path(__file__).parent / "data" / "output"

def run():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.read_excel(INPUT_DIR / "ESTIMATE.xlsx", sheet_name="in")
    is_task = df["Type"] == "Task"
    df[is_task].to_excel(OUTPUT_DIR / "ESTIMATE_Task.xlsx", index=False)
    df[~is_task].to_excel(OUTPUT_DIR / "ESTIMATE_Issue.xlsx", index=False)
    print(f"ESTIMATE_Task.xlsx:  {is_task.sum()} rows")
    print(f"ESTIMATE_Issue.xlsx: {(~is_task).sum()} rows")

if __name__ == "__main__":
    run()
