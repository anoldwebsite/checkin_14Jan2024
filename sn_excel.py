import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def get_excel_file_path(message):
    # Create a Tkinter root window
    root = tk.Tk()
    root.withdraw()
    # Show a message to the user
    messagebox.showinfo("Select Excel File", message)
    # Open a file dialog to select the Excel file
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    return file_path


def make_list_from_excel():
    file_path = get_excel_file_path(
        "Please select an Excel file that has serial numbers of assets in the first column.")
    df = pd.read_excel(file_path)
    if df.iloc[0, 0] == "Serial number" and df.iloc[0, 1] == "Lease End Date Minimum":
        df = df.iloc[1:, :]
    return df.iloc[:, 0].drop_duplicates().tolist()


import pandas as pd


def make_list_from_excel_using_substate():
    file_path = get_excel_file_path(
        "Please select an Excel file that has serial numbers and substate of assets.")

    # Read the Excel file without specifying columns
    df = pd.read_excel(file_path)

    # Select only the desired columns
    df = df[["Serial number", "Substate"]]

    # Assuming the first row contains headers, and you want to skip it
    df = df.iloc[1:, :]

    # Replace values in the second column based on specified conditions
    df["Substate"] = df["Substate"].replace({
        "Unimaged": "INFOSYSEON_SE_STAGING",
        "Defective": "INFOSYSEON_SE_RETURN",
        "Pending disposal": "INFOSYSEON_SE_RETURN"
    })

    # Create a dictionary with serial number as key and substate as value
    sn_substate_dict = dict(zip(df["Serial number"], df["Substate"]))

    return sn_substate_dict
