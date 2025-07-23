#for single-item PBIs
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
def build_grouped_description_html(page_name, page_url, testing_account_html, remediation_list):
    return (
        "<h1>PBI Goal</h1>"
        f"<p>Update the {page_name} page's [general description of component update to be made] ...</p><br />"
        "<ul>"
        f'<li><a href="{page_url}">Reference page</a>{testing_account_html}</li>'
        f"<ul>{remediation_list}</ul>"
        "</ul><br />"
        "<p>These changes are important because [why this matters for users]</p>"
    )

def render_grouped_remediations(group_entries):
    html = "<ol>"
    for entry in group_entries:
        html += "<li>"
        html += f"<strong>Conformance Recommendation:</strong> {entry['recommendation']}<br>"
        html += f"<strong>Notes:</strong> {entry['notes']}<br>"
        html += f"<strong>Remediation:</strong> {entry['remediation']}<br>"
        if entry.get("description"):
            html += f"<strong>Additional Information:</strong> {entry['description']}<br>"
        if entry.get("resources"):
            html += "<strong>Resources:</strong><ul>" + "".join(entry["resources"]) + "</ul>"
        html += "</li>"
    html += "</ol>"
    return html


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
def build_custom_acceptance_criteria_list(list_items_html):
    return (
        f"<h2>Testing Requirements</h2>"
        f"<ul>{list_items_html}</ul>"    
    )

# Wraps a single acceptance criteria item in a <p> with a heading.
def build_custom_acceptance_criteria_paragraph(text_html):
    return (
        f"<h2>Testing Requirements</h2>"
        f"<p>{text_html}</p>"
    )

# For grouped custom acceptance criteria
def build_grouped_acceptance_criteria_html(group_entries, formatter_fn):
    # Intro and key
    html = (
        "<h2>Testing Requirements</h2>"
        "<p><em>Each numbered item below corresponds to the same numbered item "
        "in the Description section above.</em></p>"
        "<ol>"
    )

    # Loop through each grouped entry
    for entry in group_entries:
        acceptance_criteria_text = entry.get("acceptance_criteria")
        ac_link = entry.get("acceptance_criteria_link")
        if acceptance_criteria_text and acceptance_criteria_text.strip():
            inner_html = formatter_fn(acceptance_criteria_text)
            inner_html = inner_html.replace("<h2>Testing Requirements</h2>", "")
            # Append Reference link if available
            if ac_link:
                inner_html += f'<p>Reference: <a href="{ac_link}">{ac_link}</a></p>'
            html += f"<li>{inner_html}</li>"
        else:
            html += (
                "<li><ul><li>"
                "<strong>TODO</strong>: [ENTER CUSTOM TESTING REQUIREMENTS FOR THIS DESCRIPTION ITEM]"
                "</li></ul></li>"
            )


    html += "</ol>"
    return html