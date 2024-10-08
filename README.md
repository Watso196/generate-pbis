# PBI Creator Script

This Python script automates the creation of accessibility remediation Product Backlog Items (PBIs) in Azure DevOps (ADO) by reading data from an accessibility audit report Excel file. It uses the Azure DevOps REST API to create PBIs based on the findings documented in the provided Excel sheet.

## Features

- Reads data from an accessibility report Excel file to create PBIs.
- Supports hyperlinks in the 'Resources, Screen Captures, Links' column.
- Generates a structured description and acceptance criteria for each PBI.
- Automatically links created PBIs to a referenced parent feature.
- Adds the PBI link to Remediation PBI column of your downloaded Excel file.

## Requirements

- Python 3.x (installing via NPM tends to be most consistent and requires the least steps)
- Libraries: `pandas`, `openpyxl`, `requests`
- Personal Access Token (PAT) for Azure DevOps

## Installation

1. **Clone this repo**.
2. **Install the required Python libraries** using pip3:
   ```
   pip3 install pandas openpyxl requests
   ```

## Preparation

Ensure that the relevant report details are filled out:

- Report Details:
  - Page Name in row 6 should be the description of the page without the word "page" included, e.g. "Category" or "PDP"
  - The URL in row 4 should lead to the page you tested against. If multiple pages were tested please only list the main page URL here!
    - Note: If you tested a production page, the script will automatically redirect to the develop page when linking the page in your PBIs.
  - The Parent Feature ID in row 12 should contain either the number ID of the parent Feature to this PBI, or a link to that Feature. The script will extract the Feature ID if you provide a link.
  - Any Account Login in row 5 should be a URL to log in to WebstaurantStore in the develop environment.
- Findings Summary:
  - Ensure that all of your data has been added to the Findings Summary.
  - Each row should only contain one issue, unless you intend to group issues of the same WCAG failure into the same PBI.
  - If you'd like the script to skip generating a PBI for an issue, you can add a note to the PBI column describing why it should be skipped. This will cause the script to skip the row.

Download an Excel copy of your accessibility audit report.

### Authentication

If you do not already have an ADO Personal Access Token (PAT) then follow these steps:

1. (Visit ADO)[https://tfs.clarkinc.biz/DefaultCollection/_work]
2. Select your user settings menu item in the top left of the screen
3. Select "Security" from the popup menu to be taken to the Security page
4. If you haven't landed on the Personal Access Tokens page, use the left-hand sidebar to navigate to Personal Access Tokens
5. Select the "New Token" button
6. Give your token a name, and select the "Full Access" radio button under "Scope"
7. Give your token a 90 day Expiration period, or longer if you'd prefer
8. Select the "Create" button
9. WARNING: Please copy and save your PAT elsewhere, as after closing the popover that allows you to copy your PAT you will be unable to copy it from the ADO interface again
10. Copy your PAT for use when running the Python file

## Usage

1. `cd` into the `generate-pbis` directory
2. Run the script from your terminal using:

```
python3 create.py
```

2. You will be prompted to enter your ADO Personal Access Token (PAT).
3. You'll be prompted to enter the path to your downloaded Excel file. Make sure the path is valid; you will have up to three attempts.
4. Upon successful creation of PBIs, you will see confirmation messages with the corresponding PBI IDs and PBI URLs.
5. The script will also write the PBI URLs into the downloaded Excel file you pointed it to. You can then either upload the newly updated Excel file (just make sure it has an appropriate name), or copy the generated Remediation PBI column data into your online Excel file.

### Test It Out!

The folder for this script also contains a test audit file linked to a (Test Feature)[https://tfs.clarkinc.biz/DefaultCollection/Design/_workitems/edit/1195637]. Feel free to run this script against the provided audit file and check the feature for the created PBIs.

Using this test file can also help debug issues with the script should you run into any.

## Important Notes

- Ensure that your Excel file is properly structured, using the latest version of the accessibility audit report. Otherwise, this script will likely fail to find important information.
- Review the Azure DevOps API documentation for any updates or changes regarding work item creation.
- If you encounter any failures with this script, please reach out to Kalib Watson with a description of your issue and your attempted resolution methods.
