"""
main.py
This module initializes the data extraction application, sets up the main window,
loads background images, screen, and buttons, and starts the main loop.

Functions:
main(): Initializes and runs the ticket machine application.
"""

from tkinter import *
import background
import screen
import search_elements
from tkinter import ttk
import session_data


def main():
    """
    Initializes and runs the ticket machine application.

    Sets up the main window with specified size and properties, loads background images,
    creates the screen and buttons, and starts the main loop.

    Parameters:
    None

    Returns:
    None
    """

    # Start tkinter
    application = Tk()

    # Window size
    application.geometry('1020x630+0+0')

    # avoid window change
    application.resizable(False, False)

    # Screen title
    application.title('Extracción de datos')

    #Screen icon
    application.wm_iconbitmap('../img/tecnm.ico')

    # Background
    background.create_background(application)

    # Creates the object to storage data
    data = session_data.SessionData()

    # progress bar
    progress_bar = ttk.Progressbar(application, orient="horizontal", length=987, mode="determinate")
    progress_bar.place(x=13, y=465)

    # Load Screen
    console_screen = screen.create_screen(application)

    # Search File button
    search_elements.select_elements_section(application, data)

    # path to save the Excel
    search_elements.select_place_to_save_file_button(application, data)

    # data extraction button
    search_elements.start_extraction_button(application, data, console_screen, progress_bar)

    # Start the main Loop
    application.mainloop()


if __name__ == "__main__":
    main()