"""
Quick end-to-end smoke test — no Node/Inspector required.

Runs a read-only `search_ownership_records` against the live API to validate
auth + request + response parsing. Uses the doc example values by default;
override with your test data from Caterpillar.

Usage:
    uv run smoke_test.py
    uv run smoke_test.py --party TD00 --dcn 12345
    uv run smoke_test.py --party DLR1 --serial 2WS23456 --make CW1
"""

import argparse
import json

import server


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--party", default="", help="partyNumber (blank = server default/override)")
    ap.add_argument("--dcn", default="", help="filter by DCN")
    ap.add_argument("--serial", default="9303", help="filter by serial number")
    ap.add_argument("--make", default="", help="filter by make code (needs a 2nd filter)")
    args = ap.parse_args()

    # 1) Auth check
    print("1) Requesting OAuth token...")
    token = server._token_cache.get()
    print(f"   OK — token length {len(token)}\n")

    # 2) Search using the flat fields (any combination, just like the agent)
    print(f"2) search_ownership_records(party={args.party!r}, dcn={args.dcn!r}, "
          f"serial_number={args.serial!r}, make_code={args.make!r})")
    result = server.search_ownership_records(
        party_number=args.party,
        dcn=args.dcn,
        serial_number=args.serial,
        make_code=args.make,
    )
    print(json.dumps(result, indent=2, default=str))

    print()
    if result.get("ok"):
        recs = (result.get("data") or {}).get("ownershipRecords", [])
        print(f"   RESULT: ok — {len(recs)} record(s). tracking_id={result.get('tracking_id')}")
    else:
        print(f"   RESULT: API returned {result.get('status_code')} — "
              f"{result.get('error')}. (Auth/transport worked; this is an API-level response.)")


if __name__ == "__main__":
    main()
