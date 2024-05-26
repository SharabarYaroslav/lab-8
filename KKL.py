import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import openpyxl

# Create a dictionary to store product information
product_info = {}

# Function to retrieve product information
def get_product_info(product):
    return product_info.get(product.lower())

# Function to handle button click event
def on_button_click():
    product = product_entry.get().strip().lower()  
    product_data = get_product_info(product)

    if product_data is not None:
        calories, photo_path, description = product_data
        result_label.config(text=f"{product.capitalize()} має {calories} калорій.", style='My.TLabel')
        
        # Display product photo
        product_image = Image.open(photo_path)
        product_image = product_image.resize((250, 250))  
        product_photo = ImageTk.PhotoImage(product_image)
        product_label.config(image=product_photo)
        product_label.image = product_photo

        # Display description
        description_text.config(state=tk.NORMAL)  # Enable editing
        description_text.delete('1.0', tk.END)    # Clear existing text
        description_text.insert(tk.END, description)
        description_text.config(state=tk.DISABLED)  # Disable editing
    else:
        result_label.config(text=f"На жаль, ми не маємо інформації про {product}.", style='My.TLabel')
        # Reset product photo and description
        product_label.config(image=None)
        description_text.config(state=tk.NORMAL)  # Enable editing
        description_text.delete('1.0', tk.END)    # Clear existing text
        description_text.config(state=tk.DISABLED)  # Disable editing

# Load data from Excel file
workbook = openpyxl.load_workbook('products.xlsx')
sheet = workbook.active

# Read data from Excel file and add it to the dictionary
for row in sheet.iter_rows(values_only=True):
    product = row[0].strip().lower()  
    calories = row[1]
    photo_path = row[2]
    description = row[3]
    product_info[product] = (calories, photo_path, description)

# Create the main window
root = tk.Tk()
root.geometry('1024x1024')

# Load background image
bg_image = Image.open("background.jpg")
bg_image = bg_image.resize((1024, 1024))
bg_photo = ImageTk.PhotoImage(bg_image)

# Create a Label for the background image
background_label = ttk.Label(root, image=bg_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Create labels, entry, and button widgets
product_label = ttk.Label(root, style='My.TLabel')
product_label.pack(pady=10)

result_label = ttk.Label(root, text="Введіть назву продукту", style='My.TLabel')
result_label.pack(pady=10)

product_entry = ttk.Entry(root, width=40, style='My.TEntry')
product_entry.pack(pady=5)

button = ttk.Button(root, text="Отримати інформацію", command=on_button_click, style='My.TButton')
button.pack(pady=5)

# Create a frame with a scrollbar for the product description
description_frame = ttk.Frame(root)
description_frame.pack(pady=10)

description_text = tk.Text(description_frame, width=58, height=10, wrap=tk.WORD, state=tk.DISABLED)
description_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

description_scrollbar = ttk.Scrollbar(description_frame, orient=tk.VERTICAL, command=description_text.yview)
description_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

description_text.config(yscrollcommand=description_scrollbar.set)

# Style the widgets
style = ttk.Style()
style.theme_create('My', parent='clam', settings={
    'TLabel': {'configure': {'font': ('montserrat', 18)}},
    'TEntry': {'configure': {'font': ('montserrat', 18), 'padding': 10, 'background': 'none'}},
    'TButton': {'configure': {'font': ('montserrat', 18), 'padding': 15, 'background': 'darkblue', 'foreground': 'white'}},
})
style.theme_use('My')

# Start the Tkinter main loop
root.mainloop()
