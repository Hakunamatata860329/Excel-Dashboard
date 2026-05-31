import argparse
import split_estimate
import build_estimate_dashboard

STEPS = {
    "data-category":      split_estimate.run,
    "estimate-dashboard": build_estimate_dashboard.run,
}

ALL_STEPS = ["data-category", "estimate-dashboard"]

def main():
    parser = argparse.ArgumentParser(description="Excel Dashboard Pipeline")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--step", nargs="+", choices=STEPS.keys(), metavar="STEP")
    group.add_argument("--all", action="store_true")
    args = parser.parse_args()

    if args.all:
        steps = ALL_STEPS
    elif args.step:
        steps = args.step
    else:
        steps = ["data-category"]

    for step in steps:
        print(f"[{step}]")
        STEPS[step]()

if __name__ == "__main__":
    main()
