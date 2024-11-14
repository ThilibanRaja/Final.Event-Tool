import os
import shutil
import sys
from jinja2 import Template

# Function to get the absolute path for resources, works for dev and bundled PyInstaller executables
def resource_path(relative_path):
    """Get the absolute path to a resource, works for dev and PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # During development, use the current working directory
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Function to update HTML template with user data
def update_html_template(event_name, date, location, registration_count, html_path):
    # Use resource_path to get the path for the HTML template
    template_path = resource_path(os.path.join("template", "demo.html"))
    
    try:
        with open(template_path, "r") as template_file:
            html_content = template_file.read()

        # Use Jinja2 to replace placeholders with user inputs
        template = Template(html_content)
        html_content = template.render(
            event_name=event_name, date=date, location=location, registration_count=registration_count
        )

        # Write the updated content to the HTML file
        with open(html_path, "w") as html_file:
            html_file.write(html_content)

    except Exception as e:
        raise Exception(f"Failed to create HTML file: {str(e)}")

# Function to copy and rename the Excel templates
def copy_excel_templates(event_name, event_folder):
    # Use resource_path to get the path for the Excel templates
    excel_template_folder = resource_path(os.path.join("excel_templates"))

    # List of template filenames
    template_files = [
        "Raw mail.xlsx",              # The first template
        "Demo.xlsx",                  # The second template
        "Demo-uPLOAD SHEET.xlsx"      # The third template
    ]
    
    # Corrected Excel template names to avoid redundancy
    renamed_files = [
        "Raw mail.xlsx",              # No change in the name for "Raw mail"
        f"{event_name}.xlsx",         # Change "Demo" to event name
        f"{event_name}-uPLOAD SHEET.xlsx"  # Change "Demo" to event name
    ]

    try:
        # Copy each Excel file and rename it according to the event name
        for i in range(len(template_files)):
            src = os.path.join(excel_template_folder, template_files[i])
            dest = os.path.join(event_folder, renamed_files[i])
            shutil.copy(src, dest)

    except Exception as e:
        raise Exception(f"Failed to copy Excel templates: {str(e)}")
