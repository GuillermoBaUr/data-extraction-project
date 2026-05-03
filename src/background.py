import tkinter as tk
from PIL import Image, ImageTk
from pathlib import Path

def create_window_with_two_colors(top_color, bottom_color, application, window_width = 1020, window_height = 630):

    # create a canvas
    canvas = tk.Canvas(application, width=window_width, height=window_height)
    canvas.pack()

    # calculate height
    top_height = window_height // 4

    # Draw the first Rectangle
    canvas.create_rectangle(0, 0, window_width, top_height, fill=top_color, outline=top_color)

    # Draw the second rectangle
    canvas.create_rectangle(0, top_height, window_width, window_height, fill=bottom_color, outline=bottom_color)


def load_tecnm_image(application, x_position):
    """
    loads the tecnm image in certain position.

    Parameters:
    application (Tk): The main Tkinter application.
    x_position (int): Horizontal Position.

    Returns:
    None
    """

    # Loading the image's path
    route = Path(r"C:\Users\guill\Desktop\Courses\Project\Tracking course management\img\tecnm.png")

    # Opening the path
    tecnm_image = Image.open(route)

    scale_image = tecnm_image.resize((100, 100), Image.LANCZOS)

    # Loading image
    tecnm_photo = ImageTk.PhotoImage(scale_image)

    # Giving image characteristics
    tecnm_label = tk.Label(application, image=tecnm_photo, bg='#007bff')

    # Placing the image
    tecnm_label.place(x=x_position, y=10)

    # Keep a reference to the image to prevent garbage collection
    tecnm_label.image = tecnm_photo

def create_title(application):
    school_name = tk.Label(application, text="Tecnológico Nacional de México", bg='#007bff', font=("Helvetica", 30, "bold"))
    campus = tk.Label(application, text="Campus San Juan Del Río", font=("Helvetica", 16, "bold"),  bg='#007bff')
    teacher_name = tk.Label(application, text="Profesora: María Patricia Uribe Rodríguez", bg='#007bff', font=("Helvetica", 16, "bold"))
    school_name.place(x=50, y=20)
    campus.place(x=55, y=70)
    teacher_name.place(x=55, y=100)


def create_background(application):

    create_window_with_two_colors("#007bff", "#f8f9fa", application)
    load_tecnm_image(application, 900)
    create_title(application)