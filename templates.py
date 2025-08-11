from helpers import safe_html

def build_description_html(page_name, page_url, testing_account_html, recommendation, notes, remediation_list, resources_html):
    return (
        "Please check the following (and delete this list when you're done!):"
        "<ul>"
        "<li>Combine any easily combine-able work into one PBI, rather than making them separate work items.</li>"
        "<li>Please keep your PBI scope to an 8 effort or lower, if possible.</li>"
        "<li>Check your tags! Aside from generated tags, you should add a tag for the site that was tested, e.g. WSS</li>"
        "<li>If your PBI needs a predecessor research PBI, create one and add the tags Spike, [Page Name] Page, Ready for Refinement. Then relate it to this PBI as a Predecessor!</li>"
        "</ul>"
        "<h1>PBI Goal</h1>"
        f"<p>Update the {page_name} page's [general description of component update to be made] ...<p><br />"
        "<ul>"
        f'<li><a href="{page_url}">Reference page</a>{testing_account_html}</li>'
        f"<li>{recommendation}</li>"
        f"<li>{notes}</li>"
        f"<ul>{remediation_list}</ul>"
        "<li>Resources:<ul>"
        f"{resources_html}"
        "</ul></li></ul><br />"
        "<p>This change is important because [why this matters for users]</p>"
    )

def render_single_remediation(remediation, description=""):
    items = [f"<li>{remediation}</li>"]
    if description:
        items.append(f"<li>Additional information: {description}</li>")
    return "".join(items)


#for grouped PBIs
def build_grouped_description_html(page_name, page_url, testing_account_html, remediation_list_html):
    return (
        "<h1>PBI Goal</h1>"
        f"<p>Update the {page_name} page's [general description of component update to be made] ...</p><br />"
        f'<p><a href="{page_url}">Reference page</a>{testing_account_html}</p>'
        f"{remediation_list_html}"
        "<br />"
        "<p>These changes are important because [why this matters for users]</p>"
    )

def render_grouped_remediations(grouped_remediation_entries):
    # Render grouped Description items as "Item 1", "Item 2"
    html_output = ""
    for item_index, remediation_entry in enumerate(grouped_remediation_entries, start=1):
        html_output += f"<h3>Item {item_index}</h3>"
        html_output += "<ul>"

        if remediation_entry.get("recommendation"):
            html_output += f"<li><strong>Conformance Recommendation:</strong> {remediation_entry['recommendation']}</li>"

        if remediation_entry.get("notes"):
            html_output += f"<li><strong>Notes:</strong> {remediation_entry['notes']}</li>"

        if remediation_entry.get("remediation"):
            html_output += f"<li><strong>Remediation:</strong> {remediation_entry['remediation']}</li>"

        if remediation_entry.get("description"):
            html_output += f"<li><strong>Additional Information:</strong> {remediation_entry['description']}</li>"

        if remediation_entry.get("resources"):
            html_output += "<li><strong>Resources:</strong><ul>"
            for resource_item in remediation_entry["resources"]:
                html_output += f"<li>{resource_item}</li>"
            html_output += "</ul></li>"

        html_output += "</ul>"

    return html_output


#default acceptance criteria for non-custom AC items 
def build_acceptance_criteria_html(page_url, page_name):
    return (
        "<h2>Testing Requirements</h2>"
        "<ul><li>Keyboard</li>"
        "<li>Screen Reader</li>"
        "<li>HTML</li>"
        "<li>etc. add any other kinds of testing you might need and remove those you don't!</li></ul>"
        "<h2>Keyboard Testing</h2>"
        f'<ul><li>Visit the <a href="{page_url}">{page_name} page</a></li>'
        "<li>[List testing steps]</li></ul>"
        "<h2>Screen Reader Testing</h2>"
        f'<ul><li>Visit the <a href="{page_url}">{page_name} page</a></li>'
        "<li>[List testing steps]</li></ul>"
        "<h2>HTML</h2>"
        f'<ul><li>Visit the <a href="{page_url}">{page_name} page</a></li>'
        "<li>[List testing steps]</li></ul>"
    )

# Wraps multiple acceptance criteria steps in a <ul> with a heading.
def build_custom_acceptance_criteria_list(list_items_html, page_url, testing_account_html):
    # Always prepend the "Visit testing page" item
    return (
        f"<h2>Testing Requirements</h2>"
        "<ul>"
        f'<li>Visit <a href="{page_url}">testing page</a>{testing_account_html}</li>'
        f"{list_items_html}"
        "</ul>"
    )

# Wraps a single acceptance criteria item in a <p> with a heading.
def build_custom_acceptance_criteria_paragraph(text_html, page_url, testing_account_html):
    # Always prepend the "Visit testing page" item even for a single AC
    return (
        f"<h2>Testing Requirements</h2>"
        "<ul>"
        f'<li>Visit <a href="{page_url}">testing page</a>{testing_account_html}</li>'
        f"<li>{text_html}</li>"
        "</ul>"
    )


# For grouped custom acceptance criteria
def build_grouped_acceptance_criteria_html(group_entries, formatter_fn, page_url, testing_account_html):
    # Intro and key
    output = (
        "<h2>Testing Requirements</h2>"
        "<p><em>Each item below corresponds to the same numbered item "
        "in the Description section above.</em></p>"
    )

    # Loop through each grouped entry
    for index, entry in enumerate(group_entries, start=1):
        acceptance_criteria_text = entry.get("acceptance_criteria")
        acceptance_criteria_link = entry.get("acceptance_criteria_link")

        # Item header
        output += f"<h3>Item {index}</h3>"

        if acceptance_criteria_text and acceptance_criteria_text.strip():
            # format the custom AC
            formatted = formatter_fn(
                acceptance_criteria_text,
                page_url,
                testing_account_html
            )
            # strip the extra heading
            formatted = formatted.replace("<h2>Testing Requirements</h2>", "")

            # Append Reference link if available
            if acceptance_criteria_link:
                # choose friendly name or fall back to the URL
                ref_name    = entry.get("acceptance_criteria_name")
                display_txt = safe_html(ref_name) if ref_name else safe_html(acceptance_criteria_link)

                formatted += (
                    f'<p>Reference: '
                    f'<a href="{safe_html(acceptance_criteria_link)}">'
                    f'{display_txt}</a></p>'
                )

            output += formatted
        else:
            output += (
                "<p><strong>TODO</strong>: [ENTER CUSTOM TESTING REQUIREMENTS FOR THIS DESCRIPTION ITEM]</p>"
            )

    return output
