import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pywhatkit as pwk
import time
import threading
import re
import os

SENT_NUMBERS_FILE = "sent_numbers.txt"  # File to track sent phone numbers

# Function to normalize phone numbers (handle spaces, dashes, country codes, and Excel auto-formatting)
def normalize_phone_number(phone_number):
    phone_number = str(phone_number)  # Ensure it's a string
    
    # Remove any spaces or dashes
    phone_number = re.sub(r'[\s\-]', '', phone_number)
    
    # Handle numbers starting with '0' and prepend country code '60' (for Malaysia)
    if phone_number.startswith('0'):
        phone_number = '60' + phone_number[1:]
    
    # Handle cases where number starts with '1' (missing country code) and has 9 digits
    if phone_number.startswith('1') and len(phone_number) == 9:
        phone_number = '60' + phone_number
    
    # Ensure the phone number starts with "+"
    if not phone_number.startswith('+'):
        phone_number = f"+{phone_number}"

    return phone_number

# Function to find the phone number column dynamically by scanning rows for keywords
def find_phone_number_column(data):
    possible_keywords = ['hand phone', 'phone', 'phone number', 'phonenumber', 'number', 'mobile', 'telephone', 'telephone no.', 'telephone no']
    
    # Iterate over the rows to search for the header
    for i, row in data.iterrows():
        if row.astype(str).str.contains('|'.join(possible_keywords), case=False, na=False).any():
            # Found the row with the phone number header
            phone_column = row[row.astype(str).str.contains('|'.join(possible_keywords), case=False, na=False)].index[0]
            # Extract phone numbers starting from the next row and remove NaNs
            data['PhoneNumber'] = data[phone_column].iloc[i+1:].dropna().apply(normalize_phone_number)
            return True
    
    return False

# Function to load an Excel file and detect the phone number column
def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        global data
        
        # Load the Excel file without assuming any header row
        data = pd.read_excel(file_path, header=None)  # Load without header

        # Search for the phone number column dynamically
        if find_phone_number_column(data):
            text_box.insert(tk.END, f"Loaded file: {file_path}\n")
            # Display normalized phone numbers in the GUI
            text_box.insert(tk.END, f"Phone numbers (normalized):\n{data['PhoneNumber'].dropna().head()}\n")  # Drop NaNs when displaying
        else:
            messagebox.showerror("Error", "Could not find a valid phone number column. Please make sure the file contains a column like 'Phone Number' or 'Hand Phone'.")

# Function to load the list of previously sent numbers from a file
def load_sent_numbers():
    if os.path.exists(SENT_NUMBERS_FILE):
        with open(SENT_NUMBERS_FILE, 'r') as f:
            sent_numbers = f.read().splitlines()
        return set(sent_numbers)  # Return as a set for faster lookup
    return set()

# Function to save a phone number to the sent numbers file
def save_sent_number(phone_number):
    with open(SENT_NUMBERS_FILE, 'a') as f:
        f.write(f"{phone_number}\n")

# Function to send WhatsApp message
def send_message(phone_number, message, load_time):
    try:
        # Normalize the phone number
        phone_number = normalize_phone_number(phone_number)
        # Send WhatsApp message with a custom loading time
        pwk.sendwhatmsg_instantly(phone_number, message, load_time)
        save_sent_number(phone_number)  # Save the number to file after sending
    except Exception as e:
        print(f"Error sending message to {phone_number}: {e}")

# Function to handle the sending of messages in the background
def start_sending(data, message, load_time):
    # Remove duplicate phone numbers from the data
    valid_numbers = data['PhoneNumber'].dropna().drop_duplicates()  # Remove duplicates within the Excel data
    total_numbers = len(valid_numbers)
    
    estimated_time = total_numbers * (load_time + 5)  # load_time + 5 seconds per message
    text_box.insert(tk.END, f"Estimated time to completion: {estimated_time // 60} minutes {estimated_time % 60} seconds\n")

    # Load sent numbers
    sent_numbers = load_sent_numbers()

    for idx, (original_index, phone_number) in enumerate(valid_numbers.items()):
        # Skip if the number has already been sent a message (from sent_numbers.txt)
        if phone_number in sent_numbers:
            text_box.insert(tk.END, f"Skipping {phone_number} (already sent)\n")
            continue

        try:
            text_box.insert(tk.END, f"{idx}: {phone_number}\n")  # Display actual progress with row numbers
            send_message(phone_number, message, load_time)
        except Exception as e:
            text_box.insert(tk.END, f"Error sending message to {phone_number}: {e}\n")
        
        time.sleep(5)  # Adding 5 seconds delay between messages
        
        progress = (idx + 1) / total_numbers * 100
        remaining_time = (total_numbers - (idx + 1)) * (load_time + 5)
        text_box.insert(tk.END, f"Progress: {progress:.2f}% - Remaining time: {remaining_time // 60} minutes {remaining_time % 60} seconds\n")
        text_box.update_idletasks()
    
    text_box.insert(tk.END, "All messages sent!\n")

# Function to handle the "Send Messages" button click
def on_send_click():
    if not data.empty:
        selected_option = project_var.get()
        # Automatically construct the message based on selection
        if selected_option == "M Suites":
            message = "Hi, I'm Sabrina. Agent for M Suites. May I know if your unit is available for rent?"
        elif selected_option == "M City":
            message = "Hi, I'm Sabrina. Agent for M City. May I know if your unit is available for rent?"

        # Get the selected load time
        load_time = int(load_time_var.get())
        # Start sending messages in a separate thread to avoid freezing the UI
        threading.Thread(target=start_sending, args=(data, message, load_time)).start()
    else:
        messagebox.showerror("Error", "Please load an Excel file with phone numbers.")

# Function to create the GUI
def create_gui():
    global text_box, data, load_time_var, project_var
    data = pd.DataFrame()  # Global empty dataframe to store numbers

    # Create the main window
    root = tk.Tk()
    root.title("WhatsApp Automation Tool")

    # Button to load Excel file
    load_button = tk.Button(root, text="Load Excel", command=load_file)
    load_button.pack()

    # Project selection (M City or M Suites)
    project_label = tk.Label(root, text="Select Project:")
    project_label.pack()

    project_var = tk.StringVar(root)
    project_var.set("M Suites")  # Default selection

    project_options = ["M Suites", "M City"]
    project_menu = tk.OptionMenu(root, project_var, *project_options)
    project_menu.pack()

    # Dropdown menu for selecting the load time
    load_time_label = tk.Label(root, text="Select browser load time (in seconds):")
    load_time_label.pack()

    load_time_var = tk.StringVar(root)
    load_time_var.set("30")  # Default to 30 seconds

    load_time_options = [str(i) for i in range(10, 61, 5)]
    load_time_menu = tk.OptionMenu(root, load_time_var, *load_time_options)
    load_time_menu.pack()

    # Button to start sending the messages
    send_button = tk.Button(root, text="Send Messages", command=on_send_click)
    send_button.pack()

    # Textbox to display progress and logs
    text_box = tk.Text(root, height=10, width=50)
    text_box.pack()

    # Run the Tkinter event loop
    root.mainloop()

# Call the function to create the GUI
create_gui()
