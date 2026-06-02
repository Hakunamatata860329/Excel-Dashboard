import argparse
import split_estimate
import build_estimate_dashboard
import build_test_suite

STEPS = {
    "data-category":      split_estimate.run,
    "estimate-dashboard": build_estimate_dashboard.run,
    "test-suite":         build_test_suite.run,
}

ALL_STEPS = ["data-category", "estimate-dashboard"]

def main():
    parser = argparse.ArgumentParser(description="Excel Dashboard Pipeline")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--step", nargs="+", choices=STEPS.keys(), metavar="STEP")
    group.add_argument("--all", action="store_true")
    parser.add_argument("--model", default="W3A", help="機種名稱，供 test-suite 使用（預設 W3A）")
    args = parser.parse_args()

    if args.all:
        steps = ALL_STEPS
    elif args.step:
        steps = args.step
    else:
        steps = ["data-category"]

    for step in steps:
        print(f"[{step}]")
        if step == "test-suite":
            STEPS[step](model=args.model)
        else:
            STEPS[step]()

if __name__ == "__main__":
    main()
