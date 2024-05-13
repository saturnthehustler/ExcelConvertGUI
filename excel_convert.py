import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def select_files(title, file_types):
    """
    Function to select multiple files using filedialog.

    Args:
    - title: Title of the file dialog window.
    - file_types: List of tuples specifying file types and extensions.

    Returns:
    - List of file paths selected by the user.
    """
    file_paths = filedialog.askopenfilenames(title=title, filetypes=file_types)
    return file_paths


def validate_excel(file_path):
    """
    Function to validate if the given file is a valid Excel file.

    Args:
    - file_path: Path of the file to validate.

    Returns:
    - True if the file is a valid Excel file, False otherwise.
    """
    try:
        pd.read_excel(file_path)  # Try reading the file
        return True
    except Exception as e:
        return False


def convert_excel():
    try:
        # Disable the Convert button
        convert_button.config(state=tk.DISABLED)

        # Display "Please wait..." message
        wait_msg = tk.Toplevel()
        wait_msg.title("Please wait...")
        wait_msg.geometry("200x100")
        wait_msg_label = tk.Label(wait_msg, text="Converting files...", font=('Arial', 12))
        wait_msg_label.pack(pady=20)

        # Select source files
        source_files = select_files("Select Excel files", [("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not source_files:
            wait_msg.destroy()
            messagebox.showerror("Error", "No files selected.")
            return

        # Validate selected files
        invalid_files = [file for file in source_files if not validate_excel(file)]
        if invalid_files:
            wait_msg.destroy()
            messagebox.showerror("Error", f"The following files are not valid Excel files:\n{', '.join(invalid_files)}")
            return

        # Select target file
        target_file = filedialog.asksaveasfilename(title="Select the target Excel file", defaultextension=".xlsx",
                                                   filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

        if not target_file:
            wait_msg.destroy()
            messagebox.showerror("Error", "Target file selection cancelled.")
            return

        # Read data from source files
        dfs = [pd.read_excel(file) for file in source_files]

        # Write data to target file
        with pd.ExcelWriter(target_file) as writer:
            for i, df in enumerate(dfs, start=1):
                df.to_excel(writer, sheet_name=f'Sheet{i}')

        # Close the "Please wait..." message
        wait_msg.destroy()

        # Show success message
        messagebox.showinfo("Success",
                            f"Data has been successfully converted and written to the target file:\n{target_file}")

    except Exception as e:
        # Close the "Please wait..." message
        wait_msg.destroy() if 'wait_msg' in locals() else None
        # Show error message
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
    finally:
        # Enable the Convert button
        convert_button.config(state=tk.NORMAL)


# Create the main window
root = tk.Tk()
root.title("Excel Converter")
root.configure(bg='#f0f0f0')  # Light gray background

# Add some padding and margins for better spacing
root.geometry("400x200")
root.resizable(False, False)

# Create a frame for better organization
frame = tk.Frame(root, bg='#f0f0f0')
frame.pack(expand=True, fill=tk.BOTH)

# Add a colorful label
label = tk.Label(frame, text="Welcome to Excel Converter", bg='#4CAF50', fg='white', font=('Arial', 20, 'bold'))
label.pack(pady=20, padx=20, fill=tk.X)

# Create a button to trigger the conversion
convert_button = tk.Button(frame, text="Convert Excel", command=convert_excel, bg='#4CAF50', fg='black',
                           font=('Arial', 14, 'bold'))
convert_button.pack(pady=10, padx=20, fill=tk.X)

# Add a quit button
quit_button = tk.Button(frame, text="Quit", command=root.quit, bg='#F44336', fg='black', font=('Arial', 14, 'bold'))
quit_button.pack(pady=10, padx=20, fill=tk.X)

root.mainloop()
