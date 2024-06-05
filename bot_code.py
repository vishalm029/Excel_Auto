# # import pandas as pd
# # import tkinter as tk
# # from tkinter import filedialog, messagebox

# # def process_excel(input_file, output_file, constraint_column, constraint_value):
# #     try:
# #         # Read the input Excel file
# #         df = pd.read_excel(input_file)

# #         # Convert the constraint column to numeric (if possible)
# #         if pd.api.types.is_numeric_dtype(df[constraint_column]):
# #             constraint_value = float(constraint_value)
# #             filtered_df = df[df[constraint_column] == constraint_value]
# #         else:
# #             filtered_df = df[df[constraint_column].astype(str) == constraint_value]

# #         # Write the filtered data to the output Excel file
# #         filtered_df.to_excel(output_file, index=False)
        
# #         messagebox.showinfo("Success", "Excel file processed successfully!")
# #     except Exception as e:
# #         messagebox.showerror("Error", f"An error occurred: {e}")

# # def browse_input_file():
# #     filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
# #     input_entry.delete(0, tk.END)
# #     input_entry.insert(0, filename)

# # def browse_output_file():
# #     filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
# #     output_entry.delete(0, tk.END)
# #     output_entry.insert(0, filename)

# # def process_button_click():
# #     input_file = input_entry.get()
# #     output_file = output_entry.get()
# #     constraint_column = constraint_column_entry.get()
# #     constraint_value = constraint_value_entry.get()
    
# #     if input_file and output_file and constraint_column and constraint_value:
# #         process_excel(input_file, output_file, constraint_column, constraint_value)
# #     else:
# #         messagebox.showerror("Error", "Please fill in all fields.")

# # # Create the main window
# # root = tk.Tk()
# # root.title("Excel Bot")

# # # Input file selection
# # input_label = tk.Label(root, text="Input Excel File:")
# # input_label.pack()
# # input_entry = tk.Entry(root, width=50)
# # input_entry.pack()
# # browse_input_button = tk.Button(root, text="Browse", command=browse_input_file)
# # browse_input_button.pack()

# # # Output file selection
# # output_label = tk.Label(root, text="Output Excel File:")
# # output_label.pack()
# # output_entry = tk.Entry(root, width=50)
# # output_entry.pack()
# # browse_output_button = tk.Button(root, text="Browse", command=browse_output_file)
# # browse_output_button.pack()

# # # Constraint selection
# # constraint_column_label = tk.Label(root, text="Constraint Column:")
# # constraint_column_label.pack()
# # constraint_column_entry = tk.Entry(root, width=30)
# # constraint_column_entry.pack()

# # constraint_value_label = tk.Label(root, text="Constraint Value:")
# # constraint_value_label.pack()
# # constraint_value_entry = tk.Entry(root, width=30)
# # constraint_value_entry.pack()

# # # Process button
# # process_button = tk.Button(root, text="Process", command=process_button_click)
# # process_button.pack()

# # # Start the GUI event loop
# # root.mainloop()



###2
# import tkinter as tk
# from tkinter import filedialog, messagebox
# from tkinter import ttk
# import pandas as pd

# def select_input_file():
#     global input_file_path
#     input_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
#     if not input_file_path:
#         messagebox.showerror("Error", "No input file selected.")
#     else:
#         input_file_label.config(text=f"{input_file_path}")

# def select_output_file():
#     global output_file_path
#     output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
#     if not output_file_path:
#         messagebox.showerror("Error", "No output file selected.")
#     else:
#         output_file_label.config(text=f"{output_file_path}")

# def process_excel():
#     if not input_file_path or not output_file_path:
#         messagebox.showerror("Error", "Please select input and output files.")
#         return
    
#     df = pd.read_excel(input_file_path)
#     constraint_col = constraint_col_entry.get()
#     constraint_val = constraint_val_entry.get()
#     target_col = target_col_entry.get()
    
#     filtered_df = df[df[constraint_col] == constraint_val]
#     target_data = filtered_df[target_col]
    
#     target_data.to_excel(output_file_path, index=False)
    
#     messagebox.showinfo("Success", "Data extracted and saved successfully.")

# # Create the GUI
# root = tk.Tk()
# root.title("Excel Data Extractor")
# root.geometry("500x400")

# input_file_path = ""
# output_file_path = ""

# # Frame for input file selection
# input_frame = ttk.Frame(root)
# input_frame.pack(pady=1)

# input_head=ttk.Label(input_frame,text="Input File:")
# input_head.pack(side="left",padx=1)

# input_frame_file_label = ttk.Frame(root)
# input_frame_file_label.pack(pady=1)

# input_file_label = ttk.Label(input_frame_file_label, text="")
# input_file_label.pack(side="left", padx=10)

# input_frame_button = ttk.Frame(root)
# input_frame_button.pack(pady=1)

# input_file_button = ttk.Button(input_frame_button, text="Select Input File", command=select_input_file)
# input_file_button.pack(side="left", padx=10)

# # Frame for output file selection
# output_frame = ttk.Frame(root)
# output_frame.pack(pady=1)

# output_head=ttk.Label(output_frame,text="Output File:")
# output_head.pack(side="left",padx=1)

# output_frame_file_label = ttk.Frame(root)
# output_frame_file_label.pack(pady=1)

# output_file_label = ttk.Label(output_frame_file_label, text="")
# output_file_label.pack(side="left", padx=10)

# output_frame_button = ttk.Frame(root)
# output_frame_button.pack(pady=1)

# output_file_button = ttk.Button(output_frame_button, text="Select Ouput File", command=select_output_file)
# output_file_button.pack(side="left", padx=10)


# # Frame for user input
# input_data_frame = ttk.Frame(root)
# input_data_frame.pack(pady=10)

# constraint_col_label = ttk.Label(input_data_frame, text="Constraint Column:")
# constraint_col_label.grid(row=0, column=0, padx=10, pady=5)

# constraint_col_entry = ttk.Entry(input_data_frame)
# constraint_col_entry.grid(row=0, column=1, padx=10, pady=5)

# constraint_val_label = ttk.Label(input_data_frame, text="Constraint Value:")
# constraint_val_label.grid(row=1, column=0, padx=10, pady=5)

# constraint_val_entry = ttk.Entry(input_data_frame)
# constraint_val_entry.grid(row=1, column=1, padx=10, pady=5)

# target_col_label = ttk.Label(input_data_frame, text="Target Column:")
# target_col_label.grid(row=2, column=0, padx=10, pady=5)

# target_col_entry = ttk.Entry(input_data_frame)
# target_col_entry.grid(row=2, column=1, padx=10, pady=5)

# # Button to process data
# process_button = ttk.Button(root, text="Process Excel", command=process_excel)
# process_button.pack(pady=10)

# root.mainloop()


#3
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd

def select_input_file():
    global input_file_path
    input_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not input_file_path:
        messagebox.showerror("Error", "No input file selected.")
    else:
        input_file_label.config(text=f"Input File: {input_file_path}")

def select_output_file():
    global output_file_path
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file_path:
        messagebox.showerror("Error", "No output file selected.")
    else:
        output_file_label.config(text=f"Output File: {output_file_path}")

def process_excel():
    if not input_file_path or not output_file_path:
        messagebox.showerror("Error", "Please select input and output files.")
        return
    
    df = pd.read_excel(input_file_path)
    constraint_col = constraint_col_entry.get()
    constraint_val = constraint_val_entry.get()
    target_col = target_col_entry.get()
    
    # Check if the constraint column is numeric or string
    if pd.api.types.is_numeric_dtype(df[constraint_col]):
        constraint_val = float(constraint_val)
        filtered_df = df[df[constraint_col] == constraint_val]
    else:
        filtered_df = df[df[constraint_col].astype(str) == constraint_val]
    
    # Check if the filtered DataFrame is empty
    if filtered_df.empty:
        messagebox.showwarning("Warning", "No data found for the given constraint.")
        return
    
    target_data = filtered_df[target_col]
    
    target_data.to_excel(output_file_path, index=False)
    
    messagebox.showinfo("Success", "Data extracted and saved successfully.")

# Create the GUI
root = tk.Tk()
root.title("Excel Data Extractor")
root.geometry("600x400")

input_file_path = ""
output_file_path = ""

# Frame for input file selection
input_frame = ttk.Frame(root)
input_frame.pack(pady=10)

input_file_label = ttk.Label(input_frame, text="Input File:")
input_file_label.pack(side="left", padx=10)

input_file_button = ttk.Button(input_frame, text="Select Input File", command=select_input_file)
input_file_button.pack(side="left", padx=10)

# Frame for output file selection
output_frame = ttk.Frame(root)
output_frame.pack(pady=10)

output_file_label = ttk.Label(output_frame, text="Output File:")
output_file_label.pack(side="left", padx=10)

output_file_button = ttk.Button(output_frame, text="Select Output File", command=select_output_file)
output_file_button.pack(side="left", padx=10)

# Frame for user input
input_data_frame = ttk.Frame(root)
input_data_frame.pack(pady=10)

constraint_col_label = ttk.Label(input_data_frame, text="Constraint Column:")
constraint_col_label.grid(row=0, column=0, padx=10, pady=5)

constraint_col_entry = ttk.Entry(input_data_frame)
constraint_col_entry.grid(row=0, column=1, padx=10, pady=5)

constraint_val_label = ttk.Label(input_data_frame, text="Constraint Value:")
constraint_val_label.grid(row=1, column=0, padx=10, pady=5)

constraint_val_entry = ttk.Entry(input_data_frame)
constraint_val_entry.grid(row=1, column=1, padx=10, pady=5)

target_col_label = ttk.Label(input_data_frame, text="Target Column:")
target_col_label.grid(row=2, column=0, padx=10, pady=5)

target_col_entry = ttk.Entry(input_data_frame)
target_col_entry.grid(row=2, column=1, padx=10, pady=5)

# Button to process data
process_button = ttk.Button(root, text="Process Excel", command=process_excel)
process_button.pack(pady=10)

root.mainloop()

