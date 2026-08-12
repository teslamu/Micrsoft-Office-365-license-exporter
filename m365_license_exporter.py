#!/usr/bin/env python3
"""
M365 License Exporter
Pull user license assignments and tenant SKU counts via Microsoft Graph API.
Copyright (c) 2026 Teslamu
Licensed under the MIT License. See LICENSE file in the project root for full license information.
"""

import csv
import os
import sys
import time
from datetime import datetime

import msal
import requests

# Default to "Microsoft Graph Command Line Tools" client ID if not specified
CLIENT_ID = os.getenv("MS_GRAPH_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e")
SCOPES = ["User.Read.All", "Directory.Read.All", "Organization.Read.All"]

# Fallback mapping so script still works if Microsoft's CSV URL changes or goes down
COMMON_SKUS = {
    "6fd2c08f-b292-42b4-b573-263ba3e01e6a": "Office 365 E3",
    "c7df2707-0d3c-4737-ba30-477270b6845f": "Office 365 E5",
    "1814162d-1a83-4c57-a97d-a30406ea4d82": "Microsoft 365 E3",
    "06a5000e-42b3-484d-8449-2a256d47377a": "Microsoft 365 E5",
    "cbdc14ab-4dbd-44e3-b61f-9a9d7cd39c30": "Microsoft 365 Business Premium",
    "3b555118-da6a-4418-894f-7df1e2096870": "Microsoft 365 Business Standard",
    "f245ace4-7444-4390-86a3-832082b6b63e": "Microsoft 365 Business Basic",
    "a403ebcc-fae0-4236-9d4e-249fb54ee8d5": "Exchange Online (Plan 1)",
    "21798e7a-3742-4726-ad13-89adc9e00867": "Exchange Online (Plan 2)",
    "c5928f49-12ba-48f7-8240-f62f3244627c": "Microsoft Teams EE",
}

SKU_CSV_URL = (
    "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/"
    "Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
)


def get_all_pages(url, headers):
    """Paginates through Graph API results (@odata.nextLink)."""
    items = []
    while url:
        resp = requests.get(url, headers=headers, timeout=20)
        
        # Simple backoff if Microsoft throttles requests
        if resp.status_code == 429:
            delay = int(resp.headers.get("Retry-After", 5))
            print(f"[!] Hit Graph API throttling. Waiting {delay}s...")
            time.sleep(delay)
            continue
            
        resp.raise_for_status()
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        
    return items


def load_sku_mappings():
    """Tries downloading official MS SKU mappings, falls back to static list on error."""
    mapping = COMMON_SKUS.copy()
    try:
        r = requests.get(SKU_CSV_URL, timeout=10)
        if r.status_code == 200:
            reader = csv.DictReader(r.text.splitlines())
            for row in reader:
                guid = row.get("GUID", "").strip()
                name = row.get("Product_Display_Name", "").strip()
                if guid and name:
                    mapping[guid] = name
            print("[+] Downloaded latest SKU list from Microsoft.")
    except Exception as err:
        print(f"[*] Couldn't fetch live SKU CSV ({err}). Using fallback SKU list.")
        
    return mapping


def main():
    print("=== M365 License Exporter ===")
    user_input = input("Enter admin email or tenant domain (leave empty for default): ").strip()

    login_hint = None
    if "@" in user_input:
        tenant = user_input.split("@")[1]
        login_hint = user_input
    elif user_input:
        tenant = user_input
    else:
        tenant = "organizations"

    authority = f"https://login.microsoftonline.com/{tenant}"
    print(f"[*] Targeting authority: {authority}")

    app = msal.PublicClientApplication(CLIENT_ID, authority=authority)
    
    auth_args = {"scopes": SCOPES}
    if login_hint:
        auth_args["login_hint"] = login_hint

    print("[*] Opening browser for authentication...")
    auth_result = app.acquire_token_interactive(**auth_args)

    if "access_token" not in auth_result:
        print(f"[-] Authentication failed: {auth_result.get('error_description')}")
        return

    headers = {
        "Authorization": f"Bearer {auth_result['access_token']}",
        "Content-Type": "application/json",
    }
    print("[+] Authenticated successfully!\n")

    sku_map = load_sku_mappings()

    print("[*] Fetching tenant users...")
    users_endpoint = (
        "https://graph.microsoft.com/v1.0/users"
        "?$select=id,displayName,userPrincipalName,assignedLicenses,accountEnabled"
    )

    try:
        users = get_all_pages(users_endpoint, headers)
    except Exception as err:
        print(f"[-] Failed getting users: {err}")
        return

    rows = []
    for u in users:
        assigned = u.get("assignedLicenses", [])
        lic_names = [
            sku_map.get(lic.get("skuId"), f"Unknown SKU ({lic.get('skuId')})")
            for lic in assigned
        ]

        rows.append({
            "DisplayName": u.get("displayName", ""),
            "UserPrincipalName": u.get("userPrincipalName", ""),
            "AccountEnabled": u.get("accountEnabled", False),
            "IsLicensed": "Yes" if assigned else "No",
            "LicenseCount": len(assigned),
            "AssignedLicenses": "; ".join(lic_names) if assigned else "Unlicensed",
        })

    # Put licensed users on top, then sort alphabetically
    rows.sort(key=lambda x: (x["IsLicensed"] == "No", x["DisplayName"].lower()))

    # File export setup
    safe_tenant_name = tenant.replace(".", "_").replace("/", "")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{safe_tenant_name}_licenses_{timestamp}.csv"

    fields = ["DisplayName", "UserPrincipalName", "AccountEnabled", "IsLicensed", "LicenseCount", "AssignedLicenses"]

    try:
        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            writer.writerows(rows)
        print(f"[+] User report written to: {filename}")
    except IOError as err:
        print(f"[-] Error writing CSV: {err}")

    licensed_count = sum(1 for r in rows if r["IsLicensed"] == "Yes")
    print("\nSummary:")
    print(f" Total Users:      {len(rows)}")
    print(f" Licensed Users:   {licensed_count}")
    print(f" Unlicensed Users: {len(rows) - licensed_count}\n")

    print("[*] Fetching overall tenant license counts...")
    skus_endpoint = "https://graph.microsoft.com/v1.0/subscribedSkus"
    
    try:
        skus = get_all_pages(skus_endpoint, headers)
        
        print("\nTenant Subscription Summary:")
        print(f"{'License':<42} | {'Purchased':<10} | {'Assigned':<10}")
        print("-" * 68)

        for s in skus:
            sku_id = s.get("skuId")
            prepaid = s.get("prepaidUnits", {})
            total_purchased = prepaid.get("enabled", 0) + prepaid.get("warning", 0)
            total_assigned = s.get("consumedUnits", 0)

            if total_purchased > 0 or total_assigned > 0:
                name = sku_map.get(sku_id, f"Unknown ({sku_id})")
                if len(name) > 40:
                    name = name[:37] + "..."
                print(f"{name:<42} | {total_purchased:<10} | {total_assigned:<10}")

        print("-" * 68)
    except Exception as err:
        print(f"[-] Could not get tenant license totals: {err}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled by user.")
        sys.exit(0)
