import sys
import re
import pandas as pd
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning  # type: ignore

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)  # type: ignore

ALATION_BASE_URL = "https://alation.genmills.com"
ALATION_TOKEN = "XnfmoyyvFRIN3PTCaDZz8KINwXEk6HemB3HbhzMiErQ"
EXCEL_FILE = "policy_run_report.xlsx"
SHEET_NAME = "DeletionCandidates"
OUTPUT_FILE = "deletion_results.csv"
VERIFY_SSL = False


def get_headers():
    return {"Token": ALATION_TOKEN}


def load_candidates():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    if "dataset_rule_id" not in df.columns:
        raise ValueError("Excel must contain 'dataset_rule_id' column")

    df = df[df["dataset_rule_id"].notna()].copy()
    df["dataset_rule_id"] = df["dataset_rule_id"].astype(str).str.strip()
    df = df[df["dataset_rule_id"] != ""]
    return df


def clean_html(value):
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"<[^>]+>", "", text)
    return text.strip()


def list_policies(skip=0, limit=100):
    url = (
        f"{ALATION_BASE_URL}/integration/v2/document/"
        f"?document_hub_id=1"
        f"&folder_id=25"
        f"&deleted=false"
        f"&limit={limit}"
        f"&skip={skip}"
    )

    print(f"Calling: {url}")

    response = requests.get(
        url,
        headers=get_headers(),
        verify=VERIFY_SSL,
        timeout=60,
    )
    print(f"Response status: {response.status_code}")
    response.raise_for_status()
    return response.json()


def extract_items(payload):
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for key in ("results", "objects", "data"):
            if isinstance(payload.get(key), list):
                return payload[key]
    return []


def extract_dataset_rule_id(rule):
    custom_fields = rule.get("custom_fields", [])

    if isinstance(custom_fields, list):
        for field in custom_fields:
            if str(field.get("field_name", "")).strip() == "Data Quality Rule ID":
                return clean_html(field.get("value", ""))

    return None


def fetch_policy_map():
    policy_map = {}
    skip = 0
    page_no = 1
    max_pages = 100

    while page_no <= max_pages:
        print(f"Fetching page {page_no}, skip={skip}...")
        payload = list_policies(skip=skip, limit=100)
        items = extract_items(payload)

        if not items:
            print("No more items found.")
            break

        print(f"Fetched {len(items)} items")

        for rule in items:
            policy_id = rule.get("id")
            dataset_rule_id = extract_dataset_rule_id(rule)
            title = str(rule.get("title", "")).strip()

            if policy_id and dataset_rule_id:
                policy_map[dataset_rule_id] = {
                    "policy_id": int(policy_id),
                    "title": title,
                }

        if len(items) < 100:
            print("Last page reached.")
            break

        skip += 100
        page_no += 1

    print(f"Total matched policies loaded: {len(policy_map)}")
    return policy_map


def delete_policy(policy_id: int):
    url = (
        f"{ALATION_BASE_URL}/integration/v2/document/"
        f"?document_hub_id=1"
        f"&folder_id=25"
        f"&deleted=false"
        f"&limit=100"
        f"&skip=0"
    )

    payload = {"id": [policy_id]}

    return requests.delete(
        url,
        json=payload,
        headers=get_headers(),
        verify=VERIFY_SSL,
        timeout=60,
    )


def main():
    dry_run = "--dry-run" in sys.argv
    df = load_candidates()
    policy_map = fetch_policy_map()

    print(f"Mode: {'DRY RUN' if dry_run else 'EXECUTE'}")
    print("Target system: Alation")
    print(f"Loaded {len(df)} input records from '{SHEET_NAME}'.")
    print(f"Fetched {len(policy_map)} policies from Alation for lookup.")

    results = []

    for _, row in df.iterrows():
        dataset_rule_id = str(row["dataset_rule_id"]).strip()
        match = policy_map.get(dataset_rule_id)

        if not match:
            print(f"[NOT FOUND] dataset_rule_id={dataset_rule_id}")
            results.append(
                {
                    "dataset_rule_id": dataset_rule_id,
                    "policy_id": "",
                    "title": "",
                    "status": "not_found",
                    "response_code": "",
                    "response_text": "No matching policy found in Alation",
                }
            )
            continue

        policy_id = match["policy_id"]
        title = match["title"]

        if dry_run:
            print(
                f"[DRY RUN] Would delete policy_id={policy_id}, "
                f"dataset_rule_id={dataset_rule_id}, title={title}"
            )
            results.append(
                {
                    "dataset_rule_id": dataset_rule_id,
                    "policy_id": policy_id,
                    "title": title,
                    "status": "would_delete",
                    "response_code": "",
                    "response_text": "",
                }
            )
            continue

        try:
            response = delete_policy(policy_id)
            status = "deleted" if response.status_code in (200, 202, 204) else "failed"

            print(
                f"[{status.upper()}] policy_id={policy_id}, "
                f"dataset_rule_id={dataset_rule_id}, "
                f"status_code={response.status_code}, response={response.text}"
            )

            results.append(
                {
                    "dataset_rule_id": dataset_rule_id,
                    "policy_id": policy_id,
                    "title": title,
                    "status": status,
                    "response_code": response.status_code,
                    "response_text": response.text,
                }
            )
        except Exception as e:
            print(
                f"[ERROR] policy_id={policy_id}, "
                f"dataset_rule_id={dataset_rule_id}, error={e}"
            )
            results.append(
                {
                    "dataset_rule_id": dataset_rule_id,
                    "policy_id": policy_id,
                    "title": title,
                    "status": "error",
                    "response_code": "",
                    "response_text": str(e),
                }
            )

    pd.DataFrame(results).to_csv(OUTPUT_FILE, index=False)
    print(f"Results saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
