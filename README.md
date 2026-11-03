# SSABL New Signings Report Automation

Turns a raw SSABL export into an organized Excel report and a Slack post.
This replaces the manual pivot-table steps.

**Scope:** United States · Select T Activate · New Signings · Data Platform, from the start of the quarter through today.

## One-time setup

1. Install Python 3.10 or newer from python.org. On Windows, tick **"Add Python to PATH"** during install.
2. Open a terminal (on Windows: Command Prompt or PowerShell) in this folder and run:
   ```
   pip install -r requirements.txt
   ```
3. Put the seller map in this folder as `SellerMap.xlsx`. It needs a sheet named `SellerMap` with these columns:
   `Coverage Type ID | Coverage Type Name | AI productivity | Data Security and Planning | TSS`
   (Copy the SellerMap tab out of an old report.)

## Each time you run the report

1. Run this once to see which dates to use:
   ```
   python ssabl_automate.py --dates-only
   ```
   It prints `Pull SSABL with Start Date ... and End Date ...`
2. Pull the report from SSABL with those dates and save it here, for example as `SSABL_export.xlsx`.
3. Run:
   ```
   python ssabl_automate.py --raw SSABL_export.xlsx --seller-map SellerMap.xlsx
   ```
4. You get:
   - `SSABL_Report_MM-DD-YY.xlsx`, with Summary, By Seller, Detail and Needs Attention tabs
   - `SSABL_Report_MM-DD-YY_slack.txt`, the Slack message ready to paste

### Posting to Slack automatically (optional)
Ask a Slack admin for an **Incoming Webhook** URL for the channel, then add it to the command:
```
python ssabl_automate.py --raw SSABL_export.xlsx --seller-map SellerMap.xlsx --slack-webhook https://hooks.slack.com/services/XXX
```
You can also set it as the environment variable `SLACK_WEBHOOK_URL`.
**Never commit the webhook URL to GitHub.** Anyone who has it can post to your channel.

### Other options
- `--out my_name.xlsx`: choose the output file name
- `--today 2026-10-27`: rebuild the report as of a past date

## Maintaining it
- **New products or categories:** edit `CATEGORY_OF_PRODUCT` near the top of `ssabl_automate.py`.
- **Territory or seller changes:** update `SellerMap.xlsx`. You don't need to change any code.
- **Different filters:** edit `FILTERS` near the top of the script.
- Anything the script can't assign to a seller goes on the **Needs Attention** tab.

## Data safety
`.gitignore` blocks all `.xlsx`, `.csv` and Slack text files, so customer names, revenue figures and seller assignments never get uploaded. Keep this repository **private**.
