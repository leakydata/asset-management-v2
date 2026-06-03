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
    ap.add_argument("--party", default="TD00", help="partyNumber (dealer code)")
    ap.add_argument("--dcn", default="12345", help="filter by DCN")
    ap.add_argument("--serial", default="", help="filter by serial number")
    ap.add_argument("--make", default="", help="filter by make code (needs a 2nd filter)")
    args = ap.parse_args()

    # 1) Auth check
    print("1) Requesting OAuth token...")
    token = server._token_cache.get()
    print(f"   OK — token length {len(token)}\n")

    # 2) Build filters from whatever was provided
    filters = []
    if args.dcn:
        filters.append({"type": "stringEquals", "propertyName": "dcn", "values": [args.dcn]})
    if args.serial:
        filters.append({"type": "stringEquals", "propertyName": "serialNumber", "values": [args.serial]})
    if args.make:
        filters.append({"type": "stringEquals", "propertyName": "makeCode", "values": [args.make]})
    if not filters:
        filters.append({"type": "stringEquals", "propertyName": "dcn", "values": ["12345"]})

    print(f"2) search_ownership_records(party={args.party!r}, filters={filters})")
    result = server.search_ownership_records(party_number=args.party, filters=filters)
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
