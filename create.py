import requests
import json
import os
import pandas as pd
from requests.auth import HTTPBasicAuth
import openpyxl
import html 

# TFS configuration
ORG_URL = "https://tfs.clarkinc.biz/DefaultCollection"
PROJECT = "Design"  # The project you're targeting
API_VERSION = "6.0"  # Update to a newer version for better support

# Function to create a PBI
def create_pbi(title, description, acceptance_criteria, priority, tags, pat):
    url = f"{ORG_URL}/{PROJECT}/_apis/wit/workitems/$Product%20Backlog%20Item?api-version={API_VERSION}"

    headers = {
        "Content-Type": "application/json-patch+json"
    }

    # Body for the work item creation
    body = [
        {"op": "add", "path": "/fields/System.Title", "value": title},
        {"op": "add", "path": "/fields/System.Description", "value": description},
        {"op": "add", "path": "/fields/Microsoft.VSTS.Common.AcceptanceCriteria", "value": acceptance_criteria},
        {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority", "value": priority},
        {"op": "add", "path": "/fields/System.IterationPath", "value": "Design\\Accessibility"},
        {"op": "add", "path": "/fields/System.AreaPath", "value": "Design\\Accessibility"},
        {"op": "add", "path": "/fields/System.Tags", "value": tags}
    ]

    try:
        # Send the request to create the work item
        response = requests.post(url, headers=headers, data=json.dumps(body), auth=HTTPBasicAuth('', pat))

        if response.status_code in (200, 201):
            print(f"Successfully created PBI: {response.json()['id']} at {response.json()['_links']['html']['href']}\n")
            return response.json()['id']  # Return the ID of the created PBI
        else:
            print(f"Failed to create PBI for {title}. Status Code: {response.status_code}, Response: {response.text}\n")
            return None
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

# Function to link the created PBI to a Parent Feature
def link_pbi_to_feature(pbi_id, feature_id, pat):
    url = f"{ORG_URL}/{PROJECT}/_apis/wit/workitems/{pbi_id}?api-version={API_VERSION}"

    # Define the relation to link the PBI to the Parent Feature
    relation_body = [
        {
            "op": "add",
            "path": "/relations/-",
            "value": {
                "rel": "System.LinkTypes.Hierarchy-Reverse",
                "url": f"{ORG_URL}/_apis/wit/workItems/{feature_id}",
                "attributes": {"comment": "Linking PBI to Parent Feature"}
            }
        }
    ]

    response = requests.patch(url, headers={"Content-Type": "application/json-patch+json"}, data=json.dumps(relation_body), auth=HTTPBasicAuth('', pat))

    if response.status_code in (200, 204):
        pass
    else:
        print(f"ERROR: Failed to link PBI. Status Code: {response.status_code}, Response: {response.text}")

# Function to map priority from text to numerical value
def map_priority(priority_text):
    priority_map = {"High": 1, "Medium": 2, "Low": 3}
    return priority_map.get(priority_text, 3)  # Default to Low (3) if not found

def write_pbi_url_to_excel(workbook, summary_sheet, row_index, pbi_url):
    # Find the index of the "Remediation PBI" column
    remediation_pbi_column = None
    
    # Iterate through the columns in the first row (headers) to find the "Remediation PBI" column
    for col in range(1, summary_sheet.max_column + 1):
        cell_value = summary_sheet.cell(row=1, column=col).value
        if cell_value == "Remediation PBI":
            remediation_pbi_column = col
            break
    
    # If "Remediation PBI" column is found, write the PBI URL in the corresponding row and column
    if remediation_pbi_column:
        summary_sheet.cell(row=row_index, column=remediation_pbi_column).value = pbi_url
    else:
        print("ERROR: 'Remediation PBI' column not found in the sheet.")

# Main function to read the Excel file and create PBIs
def create_pbis_from_excel(excel_path, pat):
    try:
        # Open the workbook and sheets
        workbook = openpyxl.load_workbook(excel_path)
        report_details_sheet = workbook['Report Details']
        summary_sheet = workbook['Findings Summary']
        
        # Extract information from the 'Report Details' sheet
        page_name = report_details_sheet.cell(row=6, column=2).value
        page_url = report_details_sheet.cell(row=4, column=2).value
        feature_id = report_details_sheet.cell(row=12, column=2).value
        testing_account_cell = report_details_sheet.cell(row=5, column=2)

         # Check that the page URL points to the development environment
        if str(page_url).startswith("https://www.webstaurantstore.com"):
            # If not a development URL, replace with the development URL
            page_url_base = "https://www.dev.webstaurantstore.com"
            page_url = page_url.replace("https://www.webstaurantstore.com", page_url_base)
        
        # Check that feature_id is not a hyperlink
        if feature_id and str(feature_id).startswith("https://"):
            # If it's a hyperlink, extract the ID from the URL
            if "?" in feature_id:
                feature_id = feature_id.split("=")[-1]
            else:
                feature_id = feature_id.split("/")[-1]
        
        # Check if the testing account cell has a hyperlink (indicating it's valid)
        if testing_account_cell.hyperlink:
            testing_account_url = testing_account_cell.hyperlink.target
            testing_account_html = f'<ul><li>Log in with <a href="{html.escape(testing_account_url)}">this account</a></li></ul>'
        else:
            testing_account_html = ""  # If no hyperlink, leave it empty

        # Read the 'Findings Summary' sheet into a DataFrame
        summary = pd.read_excel(excel_path, sheet_name='Findings Summary', engine='openpyxl')

        # Get the column index of the 'Resources, Screen Captures, Links' column
        resources_column_index = summary.columns.get_loc('Resources, Screen Captures, Links') + 1  # +1 for openpyxl index

        # List to hold PBI URLs for writing after PBIs are created
        pbi_urls = []

        # Loop through each row in the DataFrame and create PBIs
        for index, row in summary.iterrows():
            # Skip rows that have a value in the 'Remediation PBI' column
            if pd.notna(row.get('Remediation PBI')):
                print(f"Skipped row {index + 2}, PBI already assigned.\n")
                continue

            # Map the columns to the corresponding PBI fields
            title = f"Remediation - {page_name} - {row['Notes'][:50]}"  # Limiting title to 50 characters

            # Escape the content to prevent HTML injection
            page_name_escaped = html.escape(page_name)
            page_url_escaped = html.escape(page_url)
            notes_escaped = html.escape(row['Notes'])
            recommendation_escaped = html.escape(row['Conformance Recommendation'])
            remediation_escaped = html.escape(row['Remediation Techniques'])

            # Build the resources list, only adding non-empty cells (ignoring whether there's a hyperlink or not)
            resources_list = []

            # Calculate the Excel row number (pandas uses 0-based, Excel uses 1-based)
            excel_row_number = index + 2

            # Get the 'Resources' cell from openpyxl
            resource_cell = summary_sheet.cell(row=excel_row_number, column=resources_column_index)

            # Check if the 'Resources' cell has a value
            if resource_cell.value:  # Only add if the cell has a value (whether it's a hyperlink or plain text)
                resource_content = html.escape(resource_cell.value)
                resources_list.append(f'<li><a href="{resource_cell.hyperlink}">{resource_content}</a></li>')
                print(f"Added resource: {resource_content}")  # Debugging log

            # Now check the columns beyond 'Resources' (starting from the next column)
            for col_index in range(resources_column_index + 1, summary_sheet.max_column + 1):
                resource_cell = summary_sheet.cell(row=excel_row_number, column=col_index)

                # Check if the cell has a value (ignoring whether there's a hyperlink)
                if resource_cell.value:  # Only add if the cell has a value
                    resource_value = html.escape(resource_cell.value)
                    resources_list.append(f"<li>{resource_value}</li>")
                    print(f"Added resource: {resource_value}")  # Debugging log

            # Only generate the <ul> block if there's actual resource content
            resources_html = "".join(resources_list) if resources_list else ""
            # Format the description with escaped content
            description = (
                "<h1>PBI Goal</h1>"
                f"<p>Update the {page_name_escaped} page's [general description of component update to be made] ...<p><br />"
                "<ul>"
                f'<li><a href="{page_url_escaped}">Reference page</a>{testing_account_html}</li>'  # Include the testing account only if it's valid
                f"<li>{recommendation_escaped}</li>"
                f"<li>{notes_escaped}</li>"
                f"<ul><li>{remediation_escaped}</li></ul>"
                "<li>Resources:<ul>"
                f"{resources_html}"
                "</ul></li></ul><br />"
                "<p>This change is important because [why this matters for users]</p>"
            )

            acceptance_criteria = (
                "<h2>Testing Requirements</h2>"
                "<ul><li>Keyboard</li>"
                "<li>Screen Reader</li>"
                "<li>HTML</li>"
                "<li>etc. add any other kinds of testing you might need and remove those you don't!</li></ul>"
                "<h2>Keyboard Testing</h2>"
                f'<ul><li>Visit the <a href="{page_url_escaped}">{page_name_escaped} page</a></li>'
                "<li>[List testing steps]</li></ul>"
                "<h2>Screen Reader Testing</h2>"
                f'<ul><li>Visit the <a href="{page_url_escaped}">{page_name_escaped} page</a></li>'
                "<li>[List testing steps]</li></ul>"
                "<h2>HTML</h2>"
                f'<ul><li>Visit the <a href="{page_url_escaped}">{page_name_escaped} page</a></li>'
                "<li>[List testing steps]</li></ul>"
            )

            priority = map_priority(row['Priority'])
            tags = f"Remediation,Accessibility,{page_name} Page"

            # Create the PBI and get its ID
            print(f"Creating PBI for row {index + 2}...")
            pbi_id = create_pbi(title, description, acceptance_criteria, priority, tags, pat)

            # If PBI was successfully created, write the PBI URL back to the Excel file
            if pbi_id:
                pbi_url = f"{ORG_URL}/_workitems/edit/{pbi_id}"
                pbi_urls.append((index + 2, pbi_url))  # Store row index and PBI URL  # Writing back to the Excel sheet

            # Link the PBI to the Parent Feature
            if pbi_id:
                link_pbi_to_feature(pbi_id, feature_id, pat)

        # Now that all PBIs are created, write PBI URLs to the Excel sheet
        for row_index, pbi_url in pbi_urls:
            write_pbi_url_to_excel(workbook, summary_sheet, row_index, pbi_url)

        print("UPDATED: All PBI URLs written into Excel file")

        # Save the workbook after writing all URLs
        workbook.save(excel_path)
        
        print("SUCCESS: PBI creation complete!")
    
    except FileNotFoundError:
        print(f"ERROR: File {excel_path} not found. Please check the path and try again.")
    except Exception as e:
        print(f"ERROR: An error occurred: {str(e)}")


if __name__ == "__main__":
    # Get Personal Access Token from user
    PAT = input("Please enter your ADO Personal Access Token (PAT): ")

    attempts = 0  # Initialize attempt counter
    max_attempts = 3  # Set maximum number of attempts

    while attempts < max_attempts:  # Limit attempts to max_attempts
        # Prompt user for the Excel file path
        excel_file_path = input("Please enter the path to the Excel file: ")

        # Check if the file exists
        if os.path.isfile(excel_file_path):
            # Create PBIs from the Excel file
            create_pbis_from_excel(excel_file_path, PAT)
            break  # Exit the loop if the file is valid
        else:
            attempts += 1  # Increment the attempt counter
            print(f"ERROR: The file '{excel_file_path}' does not exist. Please provide a valid path.")

            # Provide feedback on remaining attempts
            remaining_attempts = max_attempts - attempts
            if remaining_attempts > 0:
                print(f"You have {remaining_attempts} attempt(s) left.")
            else:
                print("You have exceeded the maximum number of attempts. Exiting the program.")
