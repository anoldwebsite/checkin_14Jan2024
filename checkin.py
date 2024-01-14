import re
import datetime
from excel_maker import make_sn_excel_file
import tkinter as tk

import laptop
# from duration import find_num_days, change_to_timestamp, make_date_obj_from_str
from hp_warranty import get_data_from_hp

# fmt = "%Y-%m-%d %H:%M:%S"  # This is the format of the lease minimum end date in service-now.com for laptops.
country = "#AK8"  # Sweden has a code #AK8 in this organization. Denmark has code #ABY
laptops = []  # Array of object where each object is a laptop of type/class Laptop which is defined in module laptop.


def clean_title(list_of_lists):
    for (index, element) in enumerate(list_of_lists):
        title = element[1]
        pattern = r"\sG\d\s"
        match = re.search(pattern, title)
        if match:
            title = title[:match.end()]
            list_of_lists[index][1] = title
        # print(list_of_lists[index][1])
    return list_of_lists


def attach_country_code_to_title(list_of_lists):
    # Make_changes_in Bezeichnung column based on the country code in Artikel column. #AK8 for SE; #ABY for DK.
    for (index, element) in enumerate(list_of_lists):
        code = element[1]
        if code.endswith("#AK8"):
            list_of_lists[index][1] = list_of_lists[index][1] + "SE"
        elif code.endswith("#ABY"):
            list_of_lists[index][1] = list_of_lists[index][1] + "DK"
        # print(list_of_lists[index][2])
    return list_of_lists


def get_laptops(laptops_obj_list):
    """
    Changes a list of laptop objects to a list of lists with strings inside each sublist.
    Example of a sub_list is:
    ["2FZ81AV#AK8", "HP EliteBook 830 G5 Notebook PC i basmodell", "1", "5CG9094M0T", "INFOSYSEON_SE_STAGING", "None", "0,01"]

    :param laptops_obj_list: A list of laptop objects
    :return: A list of lists in which each sub-list is a string representing a feature of laptop e.g, serial number, product id etc.
    """
    laptops_str_list = []
    for one_laptop in laptops_obj_list:
        laptops_str_list.append(one_laptop.display_laptop())
    return laptops_str_list


def make_status_dic(serial_lease_dic):
    """
    Makes a dictionary of status of laptops based on number of days left which is calculated
    from the Lease End Date. If a laptop has more than 180 days left in the lease, its status
    will be set to "INFOSYSEON_SE_STAGING" otherwise "INFOSYSEON_SE_RETURN".
    :param serial_lease_dic: A dictionary in which keys are serial numbers of laptops and values are lease end dates.
    :param fmt: Date format string (default is "%Y-%m-%d %H:%M:%S").
    :return: A dictionary in which serial number is a key and value is either "INFOSYSEON_SE_STAGING" or "INFOSYSEON_SE_RETURN"
    """
    s_dic = {}
    fmt = find_date_format(serial_lease_dic)
    if not fmt:
        fmt = "%Y-%m-%d %H:%M:%S"
    for serial, lease_date in serial_lease_dic.items():
        lease_date_obj = datetime.datetime.strptime(lease_date, fmt)
        ts = lease_date_obj.timestamp()
        num_days = find_num_days(ts)
        # print(num_days)
        value = "INFOSYSEON_SE_STAGING" if num_days >= 180 else "INFOSYSEON_SE_RETURN"
        s_dic[serial] = value
    return s_dic


def find_num_days(s, e=None):
    """
    Calculates the number of days between two Unix timestamps.
    :param s: Start timestamp.
    :param e: End timestamp (default is current time).
    :return: Number of days between the two timestamps.
    """
    if e is None:
        e = datetime.datetime.utcnow().timestamp()
    if s > e:
        s, e = e, s
    dt1 = datetime.datetime.fromtimestamp(s)
    dt2 = datetime.datetime.fromtimestamp(e)
    delta = dt2 - dt1
    return int(delta.days)


def find_date_format(serial_lease_dic):
    # Try to parse the first value as %Y-%m-%d %H:%M:%S
    try:
        datetime.datetime.strptime(list(serial_lease_dic.values())[0], "%Y-%m-%d %H:%M:%S")
        return "%Y-%m-%d %H:%M:%S"
    except ValueError:
        pass

    # Try to parse the first value as %Y-%m-%d
    try:
        datetime.datetime.strptime(list(serial_lease_dic.values())[0], "%Y-%m-%d")
        return "%Y-%m-%d"
    except ValueError:
        pass

    # If neither format works, return None
    return None


import laptop  # Assuming laptop module contains the Laptop class


def make_laptops(substate_dict, machines_list):
    machines = []

    for obj in machines_list:
        serial_number = obj['serial']
        substate = substate_dict.get(serial_number,
                                     "Unknown")  # Default to "Unknown" if serial number not found in substate_dict
        machines.append(
            laptop.Laptop.customize_laptop(
                obj['pid'],
                country,
                obj['model'],
                serial_number,
                substate
            )
        )

    return machines


def make_laptops2(serial_lease_dic, machines_list):
    machines = []
    # status_dic has serial number of a laptop as key. The value is either "INFOSYSEON_SE_STAGING" or "INFOSYSEON_SE_RETURN".
    status_dic = make_status_dic(serial_lease_dic)
    # ["2FZ81AV#AK8", "HP EliteBook 830 G5 Notebook PC i basmodell", "1", "5CG9094M0T", "INFOSYSEON_SE_STAGING", "None", "0,01"]
    for obj in machines_list:
        machines.append(
            laptop.Laptop.customize_laptop(
                obj['pid'],
                country,
                obj['model'],
                obj['serial'],
                status_dic[obj['serial']]))
    return machines


def checkin_assets(serial_numbers, substate_dict):
    spt_list = get_data_from_hp(serial_numbers)  # serial_pid_title_list
    if not spt_list:  # check if the list is empty
        root_local = tk.Tk()
        root_local.withdraw()
        tk.messagebox.showerror(title="Error", message="HP website did not return the required data! Invalid Input!")
        root_local.destroy()
        return None

    print(f"Data scraped successfully for: {serial_numbers}")

    machines_list = make_laptops(substate_dict, spt_list)
    list_of_lists = get_laptops(machines_list)
    list_of_lists = clean_title(list_of_lists)
    list_of_lists = attach_country_code_to_title(list_of_lists)

    return list_of_lists


def checkin_assets2(sl, sld):
    spt_list = get_data_from_hp(sl)  # serial_pid_title_list
    if not spt_list:  # check if the list is empty
        root_local = tk.Tk()
        root_local.withdraw()
        tk.messagebox.showerror(title="Error", message="HP website did not return the required data! Invalid Input!")
        root_local.destroy()
        return None
    print(f"Data scrapped successfully for: {sl}")
    machines_list = make_laptops(sld, spt_list)
    list_of_lists = get_laptops(machines_list)
    list_of_lists = (clean_title(list_of_lists))
    list_of_lists = attach_country_code_to_title(list_of_lists)
    # print(list_of_lists)
    return list_of_lists
