import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
from PyPDF2 import PdfReader
import win32api
import threading
from striprtf.striprtf import rtf_to_text
import itertools

# Function to handle resource paths in both script and bundled executable modes
def resource(relative_path):
    base_path = getattr(
        sys,
        '_MEIPASS',
        os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Function to open folder selection dialog
def browse_folder():
    if not search_all_var.get():
        input_folder = filedialog.askdirectory()
        if input_folder:
            folder_path.delete(0, tk.END)
            folder_path.insert(0, input_folder)
        update_icon()
    else:
        folder_path.delete(0, tk.END)
        folder_path.insert(0, "Search entire computer")
        icon_label.config(image=icon2)

# Function to enable/disable folder selection based on "Search entire computer" checkbox
def lock_folder_selection():
    if search_all_var.get():
        folder_path.config(state='disabled')
        browse_button.config(state='disabled')
        icon_label.config(image=icon2)
    else:
        folder_path.config(state='normal')
        browse_button.config(state='normal')
        update_icon()

# Function to update the folder icon based on selection
def update_icon():
    if folder_path.get() and folder_path.get() != "Search entire computer":
        icon_label.config(image=icon2)
    else:
        icon_label.config(image=icon1)

# Function to start the search process in a separate thread
def search_files_thread():
    search_terms = [keyword_entry.get()]
    search_terms = [term.lower() for term in search_terms if term]
    folder_path_str = folder_path.get()
    
    if not search_terms:
        messagebox.showwarning("Warning", "Please enter at least one search term.")
        return
    
    if not folder_path_str and not search_all_var.get():
        messagebox.showwarning("Warning", "Please select a folder or check 'Search entire computer'.")
        return
    
    # Disable the search button during search
    search_button.config(state='disabled')
    
    # Start the search thread
    search_thread = threading.Thread(target=search_files)
    search_thread.start()
    
    # Start the animation
    animate_dots(search_thread)

# Function to animate the "Searching..." text
def animate_dots(search_thread):
    dots = ['.', '..', '...']
    current_dot = 0

    def update():
        nonlocal current_dot
        if search_thread.is_alive():
            result_label.config(text=f"Searching{dots[current_dot]}", foreground='green')
            current_dot = (current_dot + 1) % len(dots)
            window.after(500, update)
        else:
            # Search is complete, stop animation and update results
            window.after(0, update_result_label, search_thread)

    update()  # Start the animation loop

# Main search function
def search_files():
    global search_results
    search_terms = [keyword_entry.get()]
    search_terms = [term.lower() for term in search_terms if term]
    folder_path_str = folder_path.get()

    file_list.delete(0, tk.END)  # Clear previous search results
    search_results = []
    total_files = 0

    # Calculate total files for progress bar maximum
    if search_all_var.get():
        drives = win32api.GetLogicalDriveStrings()
        drives = drives.split('\000')[:-1]  # Remove the empty string at the end
        for drive in drives:
            for root, dirs, files in os.walk(drive):
                total_files += len(files)
    else:
        for root, dirs, files in os.walk(folder_path_str):
            total_files += len(files)
    
    progress_bar["maximum"] = total_files  # Set progress bar maximum
    
    current_progress = 0  # Track progress

    # Search logic across all drives or selected folder
    if search_all_var.get():
        for drive in drives:
            for root, dirs, files in os.walk(drive):
                for file in files:
                    file_path = os.path.join(root, file)
                    if search_file(file_path, search_terms):
                        search_results.append(file_path)
                    current_progress += 1
                    progress_bar["value"] = current_progress  # Update progress bar
                    window.update_idletasks()
    else:
        for root, dirs, files in os.walk(folder_path_str):
            for file in files:
                file_path = os.path.join(root, file)
                if search_file(file_path, search_terms):
                    search_results.append(file_path)
                current_progress += 1
                progress_bar["value"] = current_progress  # Update progress bar
                window.update_idletasks()

# Function to search individual files
def search_file(file_path, search_terms):
    _, ext = os.path.splitext(file_path)
    file_name = os.path.basename(file_path).lower()
    
    # Check if the search term is in the file name (for all file types when option is selected)
    if include_non_text_var.get() and any(term in file_name for term in search_terms):
        return True
    
    # For text-based files, always search the content
    if ext.lower() in ['.txt', '.docx', '.pdf', '.rtf', '.doc']:
        return search_text_file(file_path, search_terms)
    
    return False

# Function to search text-based files
def search_text_file(file_path, search_terms):
    _, ext = os.path.splitext(file_path)
    try:
        if ext.lower() == '.pdf':
            with open(file_path, 'rb') as f:
                reader = PdfReader(f)
                for page in reader.pages:
                    text = page.extract_text().lower()
                    if any(term in text for term in search_terms):
                        return True
        elif ext.lower() in ['.docx', '.doc']:
            import docx
            doc = docx.Document(file_path)
            full_text = '\n'.join([para.text for para in doc.paragrapaths]).lower()
            if any(term in full_text for term in search_terms):
                return True
        elif ext.lower() == '.rtf':
            with open(file_path, 'r') as rtf_file:
                rtf_content = rtf_file.read()
                text_content = rtf_to_text(rtf_content).lower()
                if any(term in text_content for term in search_terms):
                    return True
        else:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read().lower()
                if any(term in text for term in search_terms):
                    return True
    except Exception as e:
        print(f"Error reading file: {file_path} - {e}")
    return False

# Function to open selected file
def open_file(event):
    selected_item = file_list.curselection()
    if selected_item:
        file_path = file_list.get(selected_item[0])
        if os.path.isfile(file_path):
            os.startfile(file_path)
        else:
            messagebox.showwarning("Warning", "The selected item is not a file.")

# Function to open folder containing selected file
def open_folder_for_selected_file():
    selected_item = file_list.curselection()
    if selected_item:
        file_path = file_list.get(selected_item[0])
        folder = os.path.dirname(file_path)
        os.startfile(folder)

# Function to handle right-click context menu
def on_right_click(event):
    if file_list.curselection():
        context_menu.post(event.x_root, event.y_root)

# Function to update result label after search completion
def update_result_label(search_thread):
    # Re-enable the search button
    search_button.config(state='normal')

    if search_results:
        for file_path in search_results:
            file_list.insert(tk.END, file_path)  # Add matching files to listbox
        result_label.config(text=f"{len(search_results)} matching files were found", foreground='green')
        height = frame.winfo_height()
        file_list.config(height=height // 2)
    else:
        result_label.config(text="No files match the entered data", foreground='red')

    # Reset progress bar
    progress_bar["value"] = 0

# Create main window
window = ctk.CTk()
window.title('Advanced Search')
ctk.set_appearance_mode("dark") 

# Set the window icon
icon_path = resource("searching.ico")  # Use the resource function to handle paths
window.iconbitmap(icon_path)

window.geometry("550x550")  
window.resizable(False, False)

# Create main frame
frame = ttk.Frame(window)
frame.pack(expand=True, fill='both', padx=20, pady=20)

label_padding = {'padx': 10, 'pady': 15, 'sticky': 'w'}
entry_padding = {'padx': 10, 'pady': 15, 'sticky': 'ew'}

# Load emoji images
telephone_image_path = resource("keyword.png")
dossier_image_path = resource("dossier.png")

telephone_image = tk.PhotoImage(file=telephone_image_path).subsample(10, 10)
dossier_image = tk.PhotoImage(file=dossier_image_path).subsample(10, 10)

# Create labels to display emoji images
telephone_label = ttk.Label(frame, image=telephone_image)
telephone_label.grid(row=0, column=0, **label_padding)

dossier_label = ttk.Label(frame, image=dossier_image)
dossier_label.grid(row=1, column=0, **label_padding)

# Create and place UI elements
ttk.Label(frame, text='Keyword:', font=('Arial', 18)).grid(row=0, column=0, padx=75, pady=15, sticky='w')
keyword_entry = ttk.Entry(frame, font=('Arial', 18))
keyword_entry.grid(row=0, column=1)

ttk.Label(frame, text='Select\na folder:', font=('Arial',18), wraplength=140).grid(row=1, column=0, padx=75, pady=15, sticky='w')
folder_path = ttk.Entry(frame, font=('Arial', 18))
folder_path.grid(row=1, column=1)

search_all_var = tk.BooleanVar()
search_all_checkbox_style = ttk.Style()
search_all_checkbox_style.configure("Search.TCheckbutton", font=('Arial', 16))

search_all_checkbox = ttk.Checkbutton(frame, text="Search entire computer", variable=search_all_var, command=lock_folder_selection, style="Search.TCheckbutton")
search_all_checkbox.grid(row=2, column=1, columnspan=2, **label_padding)

include_non_text_var = tk.BooleanVar()
include_non_text_checkbox = ttk.Checkbutton(frame, text="Include non-text files", variable=include_non_text_var, style="Search.TCheckbutton")
include_non_text_checkbox.grid(row=3, column=1, columnspan=2, **label_padding)

browse_button_style = ttk.Style()
browse_button_style.configure("Browse.TButton", font=('Arial', 18))

# Load folder icons
icon1_path = resource("icon1.png")
icon2_path = resource("icon2.png")

icon1 = tk.PhotoImage(file=icon1_path).subsample(2, 2)
icon2 = tk.PhotoImage(file=icon2_path).subsample(2, 2)

icon_label = ttk.Label(frame, image=icon1)
icon_label.grid(row=0, column=2)

browse_button = ttk.Button(frame, text="Browse...", command=browse_folder, style="Browse.TButton")
browse_button.grid(row=1, column=2, **label_padding)

search_button_style = ttk.Style()
search_button_style.configure("Search.TButton", font=('Arial', 18))

search_button = ttk.Button(frame, text="Search", command=search_files_thread, style="Search.TButton")
search_button.grid(row=5, column=1, sticky='ew', pady=30)

# Create and configure file list
file_list = tk.Listbox(frame, font=('Arial', 13), selectmode=tk.SINGLE)
file_list.grid(row=4, column=0, columnspan=3, sticky='ew', padx=(10,0))
file_list.config(height=10)  # Doubled height

scrollbar = ttk.Scrollbar(frame, orient="vertical", command=file_list.yview)
scrollbar.grid(row=4, column=3, sticky='ns')
file_list.config(yscrollcommand=scrollbar.set)

# Bind events to file list
file_list.bind('<Double-Button-1>', open_file)
file_list.bind('<Button-3>', on_right_click)

# Create result label and progress bar
result_label = ttk.Label(frame, text='', font=('Arial', 16))
result_label.grid(row=6, column=0, columnspan=4, pady=(10, 0))

progress_bar = ttk.Progressbar(frame, orient='horizontal', mode='determinate')
progress_bar.grid(row=7, column=0, columnspan=4, padx=(0,0), pady=20, sticky='ew')

# Configure grid layout
frame.columnconfigure(1, weight=1)
frame.rowconfigure(4, weight=1)

# Initialize the context menu
context_menu = tk.Menu(window, tearoff=0, font=('Arial', 15))
context_menu.add_command(label="Open folder containing this file", command=open_folder_for_selected_file)

# Start the main event loop
window.mainloop()