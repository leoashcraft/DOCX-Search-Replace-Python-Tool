import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import configparser
import zipfile
import os
import shutil
import webbrowser

# Initialize the config parser
config = configparser.ConfigParser()
config.optionxform = str  # Preserve case of keys
config_file = 'settings.ini'

# Initialize the main window
root = tk.Tk()
root.title("Cover Letter Customization Tool")
root.geometry("600x420")
root.minsize(600, 420)
root.maxsize(600, 420)
root['background'] = '#FFE484'

style = ttk.Style(root)
style.theme_use('clam')
style.configure("Vertical.TScrollbar", troughcolor='#FFCC33', background='#FFCC33', activebackground='#FFE484')

# Define the entry widgets
entry_incoming_filepath = tk.Entry(root)
entry_outgoing_folder = tk.Entry(root)
entry_outgoing_filename = tk.Entry(root)
listbox_replacements = tk.Listbox(root)
log_text = tk.Text(root, height=10)

replacements = {}

global hotkeys_popup
hotkeys_popup = None
hotkeys = {}
        
def load_config():
    config.read(config_file)
    if 'DEFAULT' not in config:
        raise ValueError("Missing 'DEFAULT' section in settings.ini")
    if 'DEFAULT' in config:
        default = config['DEFAULT']
        entry_incoming_filepath.insert(0, default.get('incoming_filepath', ''))
        entry_outgoing_folder.insert(0, default.get('outgoing_folder', ''))
        entry_outgoing_filename.insert(0, default.get('outgoing_filename', ''))
        
        for key in hotkeys.keys():
            hotkeys[key].delete(0, tk.END)
            hotkeys[key].insert(0, default.get(key, ''))

        replacements.clear()
        for key, value in default.items():
            if key not in ['incoming_filepath', 'outgoing_folder', 'outgoing_filename'] + list(hotkeys.keys()):
                replacements[key] = value
        update_replacement_list()

def save_config():
    config['DEFAULT'] = {
        'incoming_filepath': entry_incoming_filepath.get(),
        'outgoing_folder': entry_outgoing_folder.get(),
        'outgoing_filename': entry_outgoing_filename.get(),
        **{k: v.get() for k, v in hotkeys.items()},  # Save the alt values
        **replacements
    }
    with open(config_file, 'w') as configfile:
        config.write(configfile)

def select_incoming_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])
    if file_path:
        entry_incoming_filepath.delete(0, tk.END)
        entry_incoming_filepath.insert(0, file_path)
        
def select_outgoing_folder():
    directory = filedialog.askdirectory()
    if directory:
        entry_outgoing_folder.delete(0, tk.END)
        entry_outgoing_folder.insert(0, directory)
        
def replace_text_in_file(file_path, replacements):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            filedata = file.read()

        update_log("Original file data loaded.")

        # Perform all replacements
        for placeholder, replacement in replacements.items():
            # Check and exclude alt_ keys before attempting replacement
            stripped_placeholder = placeholder.strip('XXX')
            if stripped_placeholder.startswith('alt_'):
                continue  # Skip this iteration if it's an alt_ key

            if placeholder in filedata:
                update_log(f"Replacing {stripped_placeholder} with {replacement}", color='blue')
                filedata = filedata.replace(placeholder, replacement)
            else:
                update_log(f"Placeholder {stripped_placeholder} not found in file data", color='red')

        update_log("File data modified.")

        # Write the modified content back to the file
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(filedata)

    except Exception as e:
        update_log(f"Error: {e}")

def generate_file(incoming_filepath, outgoing_folder_base, outgoing_filename, replacements):
    # Extract the folder and filename from the incoming file path
    incoming_folder = os.path.dirname(incoming_filepath)
    incoming_filename = os.path.basename(incoming_filepath)

    temp_dir = os.path.join(incoming_folder, "temp")
    temp_zip = os.path.join(incoming_folder, "temporary.zip")
    
    # Start with the base outgoing folder
    outgoing_folder_base = os.path.normpath(outgoing_folder_base)
    outgoing_folder = outgoing_folder_base

    # Use only the first two replacement values for the folder structure
    replacement_values = list(replacements.values())[:2]
    for value in replacement_values:
        outgoing_folder = os.path.join(outgoing_folder, value)

    # Create the outgoing folder if it doesn't exist
    if not os.path.exists(outgoing_folder):
        os.makedirs(outgoing_folder)

    final_path = os.path.normpath(os.path.join(outgoing_folder, outgoing_filename))
    
    # Remove temporary.zip if it exists
    if os.path.exists(temp_zip):
        os.remove(temp_zip)

    # Ensure temp directory is clean
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    try:
        with zipfile.ZipFile(os.path.join(incoming_folder, incoming_filename), 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        replace_text_in_file(document_xml_path, replacements)

        with zipfile.ZipFile(temp_zip, 'w') as zip_ref:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    zip_ref.write(file_path, os.path.relpath(file_path, temp_dir))

        # Check if the outgoing file already exists
        if os.path.exists(final_path):
            response = messagebox.askyesno("Replace File", f"The file {outgoing_filename} already exists. Do you want to replace it?")
            if not response:
                # If user chooses not to replace, clean up and exit the function
                update_log(f"File save aborted.", color='orange')
                return

        # If user chooses to replace or if the file does not exist, continue to replace it
        if os.path.exists(final_path):
            os.remove(final_path)
        os.rename(temp_zip, final_path)
        update_log(f"File saved successfully as {final_path}", color='green')
        

    except PermissionError as e:
        if e.winerror == 32:
            update_log(f"Permission Error: The file is open in another program. Please close it and try again.", color='red')
        else:
            update_log(f"Permission Error: {e}", color='red')
    except OSError as e:
        if e.winerror == 32:
            update_log(f"OS Error: The file is open in another program. Please close it and try again.", color='red')

        else:
            update_log("OS Error: {e}", color='red')

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

    finally:
        # Clean up in case of an error or if the process completes successfully
        if os.path.exists(temp_zip):
            os.remove(temp_zip)
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def add_replacement():
    key = 'XXX' + entry_key.get().strip().upper() + 'XXX'  # Automatically add 'XXX' around the key
    value = entry_value.get().strip()
    replacements[key] = value  # Store the key with 'XXX' in the dictionary
    update_replacement_list()
    save_config()

def delete_selected_replacement():
    selected = listbox_replacements.curselection()
    if selected:
        display_key = listbox_replacements.get(selected[0]).split(':')[0]
        key = 'XXX' + display_key.strip().upper() + 'XXX'  # Reconstruct the key with 'XXX'
        try:
            del replacements[key]  # Delete using the full key
            update_replacement_list()
            save_config()
        except KeyError as e:
            messagebox.showerror("Error", f"The key '{key}' was not found in the replacements.")
            
def on_replacement_selected(event):
    # Get the index of the selected replacement
    selection = listbox_replacements.curselection()
    if selection:
        index = selection[0]
        data = listbox_replacements.get(index)
        
        # Split the data by the separator (ensure this matches how you insert the data into the listbox)
        key, value = data.split(': ')
        
        # Update the entry fields, stripping the 'XXX' from the key
        entry_key.delete(0, tk.END)
        entry_key.insert(0, key.strip('XXX'))

        entry_value.delete(0, tk.END)
        entry_value.insert(0, value)

def update_replacement_list():
    listbox_replacements.delete(0, tk.END)
    for key, value in replacements.items():
        if not key.startswith('alt_'):  # Exclude alt_ keys
            display_key = key.strip('XXX')  # Removing 'XXX' from the display
            listbox_replacements.insert(tk.END, f"{display_key}: {value}")
  
# Now, set up the log Text widget with grid as well
label_value = tk.Label(root, text="Process Log:", bg='#FFE484')
label_value.grid(row=8, column=0, sticky='w', padx=5, pady=1)

# Define the log Text widget first
log_text.grid(row=9, column=0, sticky='nsew', rowspan=5, columnspan=3, padx=(5, 0), pady=1)

# Now that log_text is defined, create the Scrollbar and set it to control log_text's yview
log_scrollbar = ttk.Scrollbar(root, orient='vertical', command=log_text.yview)
log_scrollbar.grid(row=9, column=3, sticky='nsw', rowspan=5, padx=(0, 5), pady=1)

# Link the log Text widget's yscrollcommand to the log_scrollbar's set method
log_text.config(yscrollcommand=log_scrollbar.set)

def update_log(message, color='black'):
    log_text.config(state='normal')
    log_text.insert(tk.END, message + "\n", color)
    log_text.tag_config(color, foreground=color)
    log_text.config(state='disabled')
    log_text.see(tk.END)
        
def callback(url):
    webbrowser.open_new(url)
    
def open_hotkeys_popup():
    global hotkeys_popup
    
    # Close the existing popup if it's open
    if hotkeys_popup is not None and hotkeys_popup.winfo_exists():
        hotkeys_popup.destroy()

    # Create a new popup
    hotkeys_popup = tk.Toplevel(root)
    hotkeys_popup.title("Set Hotkeys")
    hotkeys_popup.geometry("300x300")  # Adjust size as needed

    hotkeys = {}  # Dictionary to store the entry widgets

    # Loop from 1 to 9
    for i in range(1, 10):
        label = tk.Label(hotkeys_popup, text=f"Alt+{i}:")
        label.grid(row=i-1, column=0)  # Adjust row index to start from 0

        entry = tk.Entry(hotkeys_popup)
        entry.grid(row=i-1, column=1)  # Adjust row index to start from 0
        hotkeys[f"alt_{i}"] = entry

    # Handle Alt+0 separately
    label = tk.Label(hotkeys_popup, text="Alt+0:")
    label.grid(row=9, column=0)  # Place Alt+0 at row 9

    entry = tk.Entry(hotkeys_popup)
    entry.grid(row=9, column=1)
    hotkeys["alt_0"] = entry

    # Button to save the values
    save_button = tk.Button(hotkeys_popup, text="Save", command=lambda: save_hotkeys(hotkeys, hotkeys_popup))
    save_button.grid(row=11, column=0, columnspan=2)

    # Load current values into the popup
    load_hotkeys_into_popup(hotkeys)
    
def save_hotkeys(hotkeys, hotkeys_popup):
    for key, entry in hotkeys.items():
        config['DEFAULT'][key] = entry.get()
    with open(config_file, 'w') as configfile:
        config.write(configfile)
    hotkeys_popup.destroy()

def load_hotkeys_into_popup(hotkeys):
    config.read(config_file)
    if 'DEFAULT' in config:
        default = config['DEFAULT']
        for key, entry in hotkeys.items():
            entry.delete(0, tk.END)
            entry.insert(0, default.get(key, ''))
          
def handle_hotkeys(event):
    # Load the current alt values from settings.ini
    config.read(config_file)
    default = config['DEFAULT']

    key = f'alt_{event.keysym}'
    alt_value = default.get(key)

    if alt_value:
        # Update the first replacement value
        if replacements:
            first_key = next(iter(replacements))
            replacements[first_key] = alt_value
            update_replacement_list()

label_replacements = tk.Label(root, text="Replacements:", bg='#FFE484')
label_replacements.grid(row=0, column=0, sticky='w', padx=5, pady=1)

listbox_replacements = tk.Listbox(root, height=5)
listbox_replacements.bind('<<ListboxSelect>>', on_replacement_selected)
listbox_replacements.grid(row=1, column=0, sticky='nsew', rowspan=6, padx=(5, 0), pady=1)

listbox_scrollbar = ttk.Scrollbar(root, orient='vertical', command=listbox_replacements.yview)
listbox_scrollbar.grid(row=1, column=1, sticky='nsw', rowspan=6, padx=(0, 5), pady=1)

listbox_replacements.config(yscrollcommand=listbox_scrollbar.set)

# UI elements for file and folder inputs
label_incoming_filepath = tk.Label(root, text="Incoming File:", bg='#FFE484')
label_incoming_filepath.grid(row=1, column=2, sticky='w', padx=5, pady=1)

entry_incoming_filepath = tk.Entry(root)
entry_incoming_filepath.grid(row=1, column=3, sticky='ew', padx=5, pady=1)

button_incoming_file = tk.Button(root, text="Browse...", bg='#FFCC33', command=select_incoming_file)
button_incoming_file.grid(row=1, column=4, padx=5, pady=1)

label_outgoing_folder = tk.Label(root, text="Outgoing Folder Path:", bg='#FFE484')
label_outgoing_folder.grid(row=2, column=2, sticky='w', padx=5, pady=1)

entry_outgoing_folder = tk.Entry(root)
entry_outgoing_folder.grid(row=2, column=3, sticky='ew', padx=5, pady=1)

button_outgoing_folder = tk.Button(root, text="Browse...", bg='#FFCC33', command=select_outgoing_folder)
button_outgoing_folder.grid(row=2, column=4, padx=5, pady=1)

label_outgoing_filename = tk.Label(root, text="Outgoing File Name:", bg='#FFE484')
label_outgoing_filename.grid(row=3, column=2, sticky='w', padx=5, pady=5)

entry_outgoing_filename = tk.Entry(root)
entry_outgoing_filename.grid(row=3, column=3, columnspan=2, sticky='ew', padx=5, pady=5)

# Load configurations
load_config()

# UI elements for replacements
label_key = tk.Label(root, text="Placeholder Text:", bg='#FFE484')
label_key.grid(row=5, column=2, sticky='w', padx=5, pady=4)

entry_key = tk.Entry(root)
entry_key.grid(row=5, column=3, columnspan=2, sticky='ew', padx=5, pady=4)

label_value = tk.Label(root, text="Replacement Text:", bg='#FFE484')
label_value.grid(row=6, column=2, sticky='w', padx=5, pady=4)

entry_value = tk.Entry(root)
entry_value.grid(row=6, column=3, columnspan=2, sticky='ew', padx=5, pady=4)

open_hotkeys_popup_button = tk.Button(root, text="Set Hotkeys", bg='#FFCC33', command=open_hotkeys_popup)
open_hotkeys_popup_button.grid(row=7, column=2, sticky='ew', padx=5, pady=1)

button_add = tk.Button(root, text="Add/Update Replacement", bg='#FFCC33', command=add_replacement)
button_add.grid(row=7, column=3, columnspan=2, sticky='ew', padx=5, pady=4)

update_replacement_list()

button_delete = tk.Button(root, text="Delete Selected Replacement", bg='#FFCC33', command=delete_selected_replacement)
button_delete.grid(row=7, column=0, sticky='ew', padx=5, pady=1)

button_generate = tk.Button(root, text="\nGenerate File\n", bg='#FFCC33', command=lambda: generate_file(
    entry_incoming_filepath.get(), 
    entry_outgoing_folder.get(), 
    entry_outgoing_filename.get(),
    replacements))
button_generate.grid(row=9, column=3, columnspan=2, ipadx=20, ipady=15, padx=5, pady=1)

label_disclaimer = tk.Label(root, text="Developed by Leo Ashcraft", fg="#d14009", bg='#FFE484', cursor="hand2")
label_disclaimer.grid(row=16, column=0, sticky='w', padx=5, pady=1)
label_disclaimer.bind("<Button-1>", lambda e: callback("https://www.leoashcraft.com/"))

# Configure the grid to expand the column and the rows where the widgets are placed
root.grid_columnconfigure(0, weight=1)  # Allows the column to expand
root.grid_rowconfigure(1, weight=1)  # Allows the row where the Listbox is placed to expand
root.grid_rowconfigure(9, weight=1)  # Allows the row where the Text widget is placed to expand

for number in range(10):  # This will loop from 0 to 9
    root.bind(f'<Alt-Key-{number}>', handle_hotkeys)
    
root.mainloop()