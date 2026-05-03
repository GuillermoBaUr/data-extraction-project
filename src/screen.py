"""
screen.py

This module contains functions to create the screen and display
information.

Functions:
create_screen(application): Creates the screen.
"""

from tkinter import Frame, FLAT, END, NORMAL, DISABLED, Text

def create_screen(application):
    """
    Creates the screen and display in the application.

    Parameters:
    application (Tk): The main Tkinter application.

    Returns:
    screen : the screen object
    """

    # Create The Frame
    screen_panel = Frame(application, bd=1, relief=FLAT, width=600, height=500)

    # Placing The Frame
    screen_panel.place(x=13, y=490)

    # Creating the screen
    screen = Text(screen_panel,
                  font=('Dosis', 10, 'bold'),
                  bd=1,
                  width=140,
                  height=7)

    # Disabling Modifications
    screen.config(state=DISABLED)

    # Placing the screen
    screen.grid(row=0,
                column=0)

    # Returning the Screen Object
    return screen


def print_text(screen, texto):

    def update_screen():
        # Activate the widget
        screen.config(state='normal')

        # Insert The text
        screen.insert('end', texto + '\n')

        # Deactivate the screen
        screen.config(state='disabled')

        # Do automatic scroll
        screen.see('end')

    screen.after(0, update_screen)
