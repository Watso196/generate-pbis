import html
import re
import pandas as pd
from openpyxl import load_workbook

def safe_html(val):
    # Safely convert a value to string if needed and escape HTML characters to prevent injection.
    return html.escape(str(val)) if pd.notna(val) else ""


def format_custom_acceptance_criteria(raw_text):
    # Converts raw Acceptance Criteria text into styled HTML.
    # Splits on '1. ', '2. ', etc. for main items.
    # Then inside each main item, splits on '*' optionally for sub‑bullets.
    # Otherwise, wraps everything in a <p>.
    
    # 1) Break into main numbered items
    potential_items = re.split(r'\d+\.\s*', raw_text)
    cleaned_items = [item.strip() for item in potential_items if item.strip()]

    if len(cleaned_items) > 1:
        list_items_html = ""

        for main_item in cleaned_items:
            # 2) Split out any "*sub‑bullet" segments
            asterisk_items = re.split(r'\*\s*', main_item)
            main_line = asterisk_items[0].strip()
            sub_items = [asterisk_item.strip() for asterisk_item in asterisk_items[1:] if asterisk_item.strip()]

            if sub_items:
                # Build nested UL
                sub_html = "".join(f"<li>{html.escape(sub_item)}</li>" for sub_item in sub_items)
                list_items_html += (
                    f"<li>{html.escape(main_line)}"
                    f"<ul>{sub_html}</ul>"
                    "</li>"
                )
            else:
                # No sub‑bullets, just render the whole thing
                list_items_html += f"<li>{html.escape(main_item)}</li>"

        from templates import build_custom_acceptance_criteria_list
        return build_custom_acceptance_criteria_list(list_items_html)
    else:
        # Single item -> paragraph
        from templates import build_custom_acceptance_criteria_paragraph
        return build_custom_acceptance_criteria_paragraph(html.escape(raw_text))


def build_resource_lookup(workbook):
    #Extracts a resource lookup from the DataLayer to provide a hyperlinked entry for the PBIs.
    sheet = workbook['DataLayer']
    resource_lookup = {}

    for row in range(2, sheet.max_row + 1):
        friendly_text = sheet.cell(row=row, column=5).value  # Column E
        url = sheet.cell(row=row, column=6).value            # Column F

        if friendly_text and url:
            resource_lookup[friendly_text.strip()] = url.strip()

    return resource_lookup
