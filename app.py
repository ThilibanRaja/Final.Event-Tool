import os
import shutil
import sys
import re
import unicodedata
from tkinter import Tk, Label, Entry, Button, StringVar, messagebox, Frame
from Source.source import update_html_template, copy_excel_templates

# Function to clean event name from hidden or special characters
def clean_event_name(name):
    # Remove leading and trailing spaces
    name = name.strip()
    # Remove any non-printable or special characters
    name = ''.join(char for char in name if char.isprintable())
    return name

# Function to forcefully replace invalid characters
def force_replace_invalid_chars(name):
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*', '\0']
    for char in invalid_chars:
        name = name.replace(char, '_')  # Replace invalid characters with underscores
    return name

# Maximum path length for Windows
MAX_PATH_LENGTH = 260

# Function to handle form submission
def submit_form():
    event_name = event_name_var.get()
    date = date_var.get()
    location = location_var.get()
    registration_count = registration_count_var.get()
    link1 = link1_var.get()
    link2 = link2_var.get()

    if not all([event_name, date, location, registration_count, link1, link2]):
        messagebox.showwarning("Input Error", "Please fill in all fields.")
        return

    # Clean and sanitize the event name
    event_name = clean_event_name(event_name)
    event_name = force_replace_invalid_chars(event_name)

    # Apply the safe folder name transformation
    event_folder = os.path.join(os.path.expanduser("~"), "Desktop", event_name)

    # Check if the path length is too long
    if len(event_folder) > MAX_PATH_LENGTH:
        messagebox.showerror("Error", "The file path is too long. Please shorten the event name.")
        return

    # Try to create the folder
    try:
        os.makedirs(event_folder, exist_ok=True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create folder. Error: {str(e)}")
        return

    # Define paths for the HTML, text file, and Excel templates
    html_path = os.path.join(event_folder, f"{event_name}-Pitch.html")
    text_path = os.path.join(event_folder, "Event Details.txt")

    # Update HTML file with provided values
    try:
        update_html_template(event_name, date, location, registration_count, html_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create HTML file. Error: {str(e)}")
        return

    # Write the event details to the text file
    try:
        with open(text_path, "w") as file:
            file.write(f"Event Name: {event_name}\n")
            file.write(f"Date: {date}\n")
            file.write(f"Location: {location}\n")
            file.write(f"Registration Count: {registration_count}\n\n")
            file.write(f"Link1: {link1}\n\n")
            file.write(f"Link2: {link2}\n\n")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create text file. Error: {str(e)}")
        return

    # Copy and rename the Excel templates
    try:
        copy_excel_templates(event_name, event_folder)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to copy Excel templates. Error: {str(e)}")
        return

    messagebox.showinfo("Success", f"Files for event '{event_name}' created successfully!")

# Set up Tkinter UI
app = Tk()
app.title("Event Folder Creator")
app.geometry("400x350")
app.configure(bg="gray")

# Frame to hold all input fields and center them
frame = Frame(app, bg="gray")
frame.place(relx=0.5, rely=0.5, anchor="center")  # Center the frame on the screen

# UI elements for user input
event_name_var = StringVar()
date_var = StringVar()
location_var = StringVar()
registration_count_var = StringVar()
link1_var = StringVar()
link2_var = StringVar()

# Use grid layout within the frame
Label(frame, text="Event Name:", bg="gray").grid(row=0, column=0, padx=10, pady=5, sticky="e")
Entry(frame, textvariable=event_name_var).grid(row=0, column=1, padx=10, pady=5)

Label(frame, text="Date:", bg="gray").grid(row=1, column=0, padx=10, pady=5, sticky="e")
Entry(frame, textvariable=date_var).grid(row=1, column=1, padx=10, pady=5)

Label(frame, text="Location:", bg="gray").grid(row=2, column=0, padx=10, pady=5, sticky="e")
Entry(frame, textvariable=location_var).grid(row=2, column=1, padx=10, pady=5)

Label(frame, text="Registration Count:", bg="gray").grid(row=3, column=0, padx=10, pady=5, sticky="e")
Entry(frame, textvariable=registration_count_var).grid(row=3, column=1, padx=10, pady=5)

Label(frame, text="Link1:", bg="gray").grid(row=4, column=0, padx=10, pady=5, sticky="e")
Entry(frame, textvariable=link1_var).grid(row=4, column=1, padx=10, pady=5)

Label(frame, text="Link2:", bg="gray").grid(row=5, column=0, padx=10, pady=5, sticky="e")
Entry(frame, textvariable=link2_var).grid(row=5, column=1, padx=10, pady=5)

Button(frame, text="Create Event Folder", command=submit_form, bg="skyblue").grid(row=6, column=0, columnspan=2, pady=20)

# Run the application
app.mainloop()
