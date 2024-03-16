import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook


# Function to search for the data
def search_data():
    id_search = entry_id.get()
    load_data(id_search)  # Pass the ID search term to load_data
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # Convert both to string for an exact match
        if str(row[0]).zfill(3) == id_search.zfill(3):
            entry_name.delete(0, tk.END)
            entry_name.insert(0, row[1])
            entry_usd.delete(0, tk.END)
            entry_usd.insert(0, str(row[2]) if row[2] is not None else "") #Update to str when search   
            entry_riel.delete(0, tk.END) 
            entry_riel.insert(0, str(row[3]) if row[3] is not None else "") #Update to str whem search
            break


# Function to clear all the entry widgets and reset the Treeview
def clear_entries():
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    entry_usd.delete(0, tk.END)
    entry_riel.delete(0, tk.END)
    load_data()  # Call load_data without arguments to reset the Treeview


# Update the update_data function to handle the Treeview selection
def update_data():
    try:
        id_search = entry_id.get()
        name_to_find = entry_name.get()  # Also get the Name 
        found = False
        for row in sheet.iter_rows(min_row=2):
            if str(row[0].value).zfill(3) == id_search.zfill(3) and row[1].value == name_to_find:  # Match Name too
                row[2].value = int(entry_usd.get())  
                row[3].value = int(entry_riel.get())
                found = True
                break  # Can break, since we found the specific row

        if found:
            wb.save('.\data.xlsx')
            print("Data updated successfully.")
            refresh_data()
        else:
            print("ID and Name combination not found. Please check the ID and Name and try again.")
    except ValueError:
        print("Please enter a valid number for USD and Riel.")
    except Exception as e:
        print(f"An error occurred: {e}")
    


# Function to fill the entry widgets with the data from the selected row
def on_tree_select(event):
    # Check if there is a selection
    if tree.selection():
        selected_item = tree.selection()[0]  # Get selected item
        # Get the values of the selected row
        selected_row = tree.item(selected_item, 'values')
        # Clear the entry fields
        clear_entries()
        # Insert the data into the entry fields
        entry_id.insert(0, selected_row[0])
        entry_name.insert(0, selected_row[1])
        entry_usd.insert(0, selected_row[2])
        entry_riel.insert(0, selected_row[3])
    else:
        print("No item selected")



# Function to load data from Excel into the Treeview
def load_data(id_search=None):
    for item in tree.get_children():
        tree.delete(item)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if id_search:
            # Check if the ID in the row starts with the entered ID
            if str(row[0]).startswith(id_search):
                tree.insert("", tk.END, values=row)
        else:
            tree.insert("", tk.END, values=row)



# Function to refresh the Treeview with updated data
def refresh_data():
    for item in tree.get_children():
        tree.delete(item)
    load_data()

    
    
# Load the workbook and select the active sheet
wb = load_workbook('.\data.xlsx')
sheet = wb.active   


# Create the main window
root = tk.Tk()
root.title("Wedding Guest Book")


# Create the entry widgets
entry_id = tk.Entry(root, width=40)
entry_id.grid(row=0, column=1)
entry_name = tk.Entry(root, width=40)
entry_name.grid(row=1, column=1)
entry_usd = tk.Entry(root, width=40)
entry_usd.grid(row=2, column=1)
entry_riel = tk.Entry(root, width=40)
entry_riel.grid(row=3, column=1)

# Create the label widgets
label_id = tk.Label(root, text="ID")
label_id.grid(row=0, column=0)
label_name = tk.Label(root, text="Name")
label_name.grid(row=1, column=0)
label_usd = tk.Label(root, text="USD")
label_usd.grid(row=2, column=0)
label_riel = tk.Label(root, text="Riel")
label_riel.grid(row=3, column=0)

# Create the button widgets
button_search = tk.Button(root, text="Search", command=search_data,width=20)
button_search.grid(row=4, column=0)
# Create the 'Clear' button and place it on the grid
button_clear = tk.Button(root, text="Clear", command=clear_entries,width=20)
button_clear.grid(row=4, column=2)
button_update = tk.Button(root, text="Update", command=update_data,width=20)
button_update.grid(row=4, column=1)

# Create the Treeview widget
tree = ttk.Treeview(root, columns=("ID", "Name", "USD", "Riel"), show='headings')
tree.heading("ID", text="ID")
tree.heading("Name", text="Name")
tree.heading("USD", text="USD")
tree.heading("Riel", text="Riel")

tree.grid(row=7, columnspan=5, sticky='nsew')
# Bind the select event of the tree to the on_tree_select function
tree.bind('<<TreeviewSelect>>', on_tree_select)
# Load data into the Treeview
load_data()

# Run the main loop
root.mainloop() #ok