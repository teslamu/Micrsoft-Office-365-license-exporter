# M365 License Exporter

A handy Python script that pulls user license assignments and tenant SKU counts via the Microsoft Graph API. 

If you've ever tried to audit M365 licenses, you know Microsoft often returns cryptic SKU GUIDs (like `c7df2707-0d3c-4737-ba30-477270b6845f`) instead of actual product names (like `Office 365 E5`). This script automatically downloads the latest SKU mappings directly from Microsoft to give you clean, human-readable data.

## Features

* **Secure Authentication:** Uses MSAL (Microsoft Authentication Library) for a secure, interactive browser login. No hardcoded passwords or client secrets required.
* **Smart SKU Translation:** Dynamically fetches Microsoft's official CSV of SKU names, with a built-in fallback just in case their link goes down.
* **Built for the Real World:** Automatically handles Graph API pagination and gracefully pauses if it hits Microsoft's rate limits (HTTP 429).
* **CSV Export:** Generates a neatly formatted CSV of all users, their account status, and their assigned licenses.
* **Terminal Summary:** Prints a quick, easy-to-read summary of your total purchased vs. assigned licenses directly in the console.

## Prerequisites

You'll need Python 3 installed, along with two external libraries: `msal` and `requests`. 

Install them using pip:
```bash
pip install msal requests
