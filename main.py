"""
https://nuitka.net/doc/user-manual.html
Download and install Python from https://www.python.org/downloads/windows

Select one of Windows x86-64 web-based installer (64 bits Python, recommended) or x86 executable (32 bits Python) installer.

Verify it’s working using command python --version.
"""
# python -m pip install nuitka
# python -m nuitka --version
# On a terminal, use the following command inside the directory where the main.py file is.
# python -m nuitka --follow-imports --standalone main.py
# pip install more-itertools
# pip install webdriver_manager
# https://stackoverflow.com/questions/25905540/importerror-no-module-named-tkinter
from sn_excel import make_list_from_excel, make_list_from_excel_using_substate, get_excel_file_path
from sn_scanner import get_unique_sn
from sn_lease_dic import read_lease_file
from checkin import checkin_assets
from excel_maker import make_checkin_excel_file, make_sn_excel_file, make_sn_substate_excel_file
import more_itertools
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.title("Check in Assets - created by Ajmal Khan")


def make_pages(sn_substate_dict):
    # Split the sn_substate_dict dictionary into chunks of 14 elements
    sn_chunks = [list(sn_substate_dict.keys())[i:i + 14] for i in range(0, len(sn_substate_dict), 14)]

    # Call checkin_assets for each chunk of 14 elements
    lst_of_lst = []
    for i, chunk in enumerate(sn_chunks):
        substate_dict_chunk = {k: sn_substate_dict[k] for k in chunk}
        l_of_l = checkin_assets(chunk, substate_dict_chunk)

        if not l_of_l:  # check if the list is empty
            print(
                "One of the serial numbers has a problem. Saving 14 serial numbers in which one or more serial numbers are problematic.")
            print(chunk)
            make_sn_substate_excel_file(substate_dict_chunk)
            if lst_of_lst and len(lst_of_lst) >= 1:
                return lst_of_lst
            else:
                return None

        lst_of_lst.extend(l_of_l)

    return lst_of_lst


def make_pages2(s_n, sn_lease_dict):
    # Split the sn list and sn_lease_dic dictionary into chunks of 14 elements
    sn_chunks = list(more_itertools.chunked(s_n, 14))  # 14 is the maximum number of labels on 1 page in our warehouse.
    sn_lease_dic_chunks = [{k: v for k, v in sn_lease_dict.items() if k in chunk} for chunk in sn_chunks]
    # Call checkin_assets for each chunk of 14 elements
    lst_of_lst = []
    for i, chunk in enumerate(sn_chunks):
        l_of_l = checkin_assets(chunk, sn_lease_dic_chunks[i])
        if not l_of_l:  # check if the list is empty
            print(
                "One of the serial numbers has problem. Saving 14 serial numbers in which one or more serial numbers are problematic.")
            print(chunk)
            make_sn_excel_file(chunk)
            return None
        lst_of_lst.extend(l_of_l)
    return lst_of_lst


def use_excel():
    sn_substate_dict = make_list_from_excel_using_substate()
    list_of_pages = make_pages(sn_substate_dict)

    if list_of_pages:
        make_checkin_excel_file(list_of_pages)
    else:
        print(
            "No data for making a complete check-in file. Instead saving both serial numbers and substates of assets.")
        make_sn_substate_excel_file(sn_substate_dict)

    root.destroy()  # Close the window after the function is executed


def use_excel2():
    sn = make_list_from_excel()
    sn_lease_dic = read_lease_file(
        get_excel_file_path(
            "Please select an Excel file that has Lease End Minimum Date for assets in its second column."))
    list_of_pages = make_pages(sn, sn_lease_dic)
    if list_of_pages:
        make_checkin_excel_file(list_of_pages)
    else:
        print("No data for making a complete check in file. Instead saving only serial numbers of assets.")
        make_sn_excel_file(sn)
    root.destroy()  # Close the window after the function is executed


def use_scanner():
    sn = get_unique_sn()
    sn_lease_dic = read_lease_file(
        get_excel_file_path(
            "Please select an Excel file that has Lease End Minimum Date for assets in its second column."))
    list_of_pages = make_pages(sn, sn_lease_dic)
    if list_of_pages:
        make_checkin_excel_file(list_of_pages)
    else:
        make_sn_excel_file(sn)
        root_local = tk.Tk()
        root_local.withdraw()
        tk.messagebox.showerror(title="Error", message="HP website did not return the required data! Invalid Input!")
        root_local.destroy()
        stop_program()

    root.destroy()  # Close the window after the function is executed


def stop_program():
    root.quit()


# Create three buttons
excel_button = tk.Button(root, text="Use Excel", command=use_excel, font=("Arial", 16), height=2, width=20)
excel_button.pack(pady=10)
scanner_button = tk.Button(root, text="Use Scanner", command=use_scanner, font=("Arial", 16), height=2, width=20)
scanner_button.pack(pady=10)
stop_button = tk.Button(root, text="Stop", command=stop_program, font=("Arial", 16), height=2, width=20)
stop_button.pack(pady=10)

# Set the size of the window
root.geometry("400x300")

root.mainloop()
