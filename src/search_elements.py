import tkinter as tk
from tkinter import ttk, filedialog
import data_extraction



def open_select_file_dialog(label, data):
    file_path = filedialog.askopenfilenames(title="Select ZIP file",
                                           filetypes=[("ZIP files", "*.zip")])

    if file_path:
        for zip_file in file_path:
            data.zip_files.append(zip_file)
        file_names = [fp.split('/')[-1] for fp in file_path]
        label.config(text=f'{len(file_names)} Selected file(s): {', '.join(file_names)}')
    else:
        label.config(text='No file selected')

def open_select_path_dialog(label, data):
    file_path = filedialog.askdirectory(title="Select Directory")

    if file_path:
        data.destination_path = file_path
        label.config(text=f'{file_path}')
    else:
        label.config(text='No file selected')


def select_elements_section(application,  data):

    # Create a label to know if a file was selected
    label = ttk.Label(application, text="No file selected", background='#f8f9fa')
    label.place(x=100, y=328)

    # Create a button to opening file searcher
    button_style = ttk.Style()
    button_style.configure("TButton", background="#f8f9fa")
    button = ttk.Button(application, text="Select ZIP File",
                                     style="TButton",
                                     command=lambda: open_select_file_dialog(label, data))
    button.place(x=13, y=325)



def start_extraction_button(application, data, console_screen, progress_bar):

    # Create a button to start data extraction
    button_style = ttk.Style()
    button_style.configure("TButton", background="#f8f9fa")
    button = ttk.Button(application, text="Extract Data",
                        style="TButton",
                        command=lambda: data_extraction.run_extract_data_in_thread(data, console_screen, progress_bar))
    button.place(x=13, y=415)

def select_place_to_save_file_button(application, data):

    # Create a label to know if a file was selected
    path_label = ttk.Label(application, text="No file selected", background='#f8f9fa')
    path_label.place(x=145, y=378)

    # Create a button to select path to save the file
    button_style = ttk.Style()
    button_style.configure("TButton", background="#f8f9fa")
    path_button = ttk.Button(application, text="Select Destination File",
                              style="TButton",
                             command=lambda: open_select_path_dialog(path_label, data))
    path_button.place(x=13, y=375)

