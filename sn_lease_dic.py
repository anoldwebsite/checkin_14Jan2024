import pandas as pd


def read_lease_file(file_name):
    """
    Converts the two columns i.e. serial numbers and lease end dates from passed Excel sheet to a dictionary.
    The keys are serial numbers taken from the first column of the Excel file received as argument.
    The values are the lease end dates taken from the second column of the Excel file received as argument.
    :param file_name: The default is lease.xlsx but the caller can pass any Excel file with two columns.
    :return: Dictionary in which the keys are serial numbers taken from the first column of the Excel file and
    values are lease end dates taken from the second column of the Excel file.
    """
    # Load the Excel file with no header row
    df = pd.read_excel(file_name, header=None)
    # Check if the first row contains headers, if yes remove it
    if df.iloc[0, 0] == "Serial number" and df.iloc[0, 1] == "Lease End Date Minimum":
        df = df.iloc[1:, :]
    # Convert the dataframe to dictionary
    serial_lease_dict = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1].astype(str)))
    # Remove possible duplicates from the dictionary
    serial_lease_dict = {k: v for k, v in serial_lease_dict.items() if str(k) != 'nan'}
    return serial_lease_dict
