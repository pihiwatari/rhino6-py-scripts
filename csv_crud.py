"""CSV CRUD for pallet project object exports.

A csv CRUD operations module to use with rhinoscriptsyntax, which
creates, read and add new data to a csv file. This files needs to
be located in the same folder as the CAM exports files in a pallet
project folder.

The end goal of this modules is to create and update a csv file with the
information of the rhino models and copy that data to an excel control
file.

Functions:
    * create_data_dict - returns an OrderedDict containing the column
    headers as key for data reference for the rhino object.
    * create_new_csv - generates a new csv file, used only for first
    time exports.
    * get_row_index - returns the last row index as reference for other
    operations, such as adding new data.
    * setup_project_data - returns an updated data dictionary with the project
    information.
"""

import rhinoscriptsyntax as rs
from collections import OrderedDict
import csv
import os

csv_headers = [
    "Row",
    "Flex Customer",
    "Site",
    "Building",
    "ID",
    "Product",
    "Requester",
    "Empty1",  # Start date
    "Empty2",  # Commit date
    "Material",
    "Designer",
    "Sizes",
    "Quantity",
    "Empty3",  # CAM date
    "Empty4",  # Programer
    "Empty5",  # Machining CNC date
    "Empty6",  # Machine
    "Empty7",  # End of machiningCNC
    "Empty8",  # Machining time
    "File Path",  # Box drive link
    "File name",
    "Special details"
]

project_columns = {
    "Flex Customer": "Flex",
    "Site": "Guad North",
    "Building": "B1",
    "ID": "DM-GDN00-000000",
    "Product": "Pallet",
    "Requester": "Requester name",
    "Designer": os.environ.get("USERNAME"),
    "File Path": "Paste Box drive link here",
}


def create_new_csv(filepath):
    # type: (str) -> None
    """Creates a new csv file.

    Used when initiating a new pallet design project, this function
    creates an empty csv only with the column csv_headers. And project
    information
    """

    # Check if a DATA file already exists in the folder
    if os.path.exists(filepath):
        print("A DATA.csv file already exists in the folder")
        return

    # Write file
    try:
        with open(filepath, "wb") as f:

            # Create csv writer
            writer = csv.writer(f)

            # Write csv_headers
            writer.writerow(csv_headers)

    except IOError:
        print("IOError")


def create_data_dict():
    # type: () -> OrderedDict
    """ Creates an OrderedDict with predefined key values.

    Returns a dictionary with predefined key values. All of the dict values
    are None by default, this is replaced in with fucntions.

    Parameters:
        None.

    Returns:
        data_dict (OrderedDict): dictionary with the predefined key values
    """

    data_dict = OrderedDict()

    for item in csv_headers:

        data_dict[item] = None

    return data_dict


def setup_project_data(data_dict):
    # type: (OrderedDict) -> OrderedDict
    """Setup recurrent project data for csv writing operations.

    Returns an updated data dictionary with the project's information,
    reducing input time from the user for repeating data.

    Parameters:
        data_dict (OrderedDict): Data dictionary to update.
        project_columns (dict[str, any]): A dict with the keys to be updated,
        and with the default values.

    Returns:
        data_dict (OrderedDict): Update dictionary with the project's
        information.
    """

    # Iterate through all the list items.
    for key in project_columns.keys():

        default_value = project_columns[key]

        # Prompt for user input on the data.
        data = rs.StringBox(
            message="Please input information for: %s." % key,
            title="Setup project data",
            default_value=default_value
        )

        if not data:
            return

        # Update dict with new value
        data_dict[key] = data

    return data_dict


def read_project_data(filepath):
    # type: (str) -> None
    """Reads project data rows from a csv file.

    This function reads project data from a csv file in the
    exports folder.

    Parameters:
        filepath (str): Path to the csv file.

    Returns:
        None.
    """

    try:
        # Open csv file.
        with open(filepath, 'r') as f:

            reader = csv.DictReader(f)

            row = next(reader)
            return row
    except IOError:
        print("""
        Error in function: read_project_data.\n 
        This csv has no data to read.""")


def get_row_index(filepath, name=None):
    # type: (str, str | None) -> int
    """Returns the last row index number in a csv file.

    Parameters:
        filepath (str): The location of the csv file to read.
        name (str | None): The name to search for in the csv file. None if omitted


    Returns:
        row_counter (int): The last index +1 of current written data 
        row in the csv. Returns -1 if the name is found.
    """

    row_counter = 0

    with open(filepath, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if row["File name"] == name:
                return -1

            row_counter += 1

    return row_counter


def add_data_row(filepath, data_dict):
    # type: (str, OrderedDict) -> None
    """Adds a new line of data to the csv file.

    Adds a single new row to a csv file with the information of a 
    RhinoGeometry object.

    Parameters:
        filepath (str): The location of the csv file to add data.
        data_dict (OrderedDict): Ordered dictionary with the data.

    Returns:
        None.
    """

    # # Get last row number in the csv
    # row_number = get_row_index(filepath)

    # # Add row number to the data dict
    # data_dict["Row"] = row_number

    try:
        with open(filepath, "ab") as f:

            writer = csv.DictWriter(f, fieldnames=csv_headers)
            writer.writerow(data_dict)
            rs.MessageBox(
                message="Succesfully saved model data:\n{}"
                        .format(data_dict),
                buttons=0,
                title="Data saved"
            )

    except IOError:
        print("""
        Error in function: add_data_row.\n
        An error ocurred while writing information to the csv.""")


def update_data_row(filepath, name, new_data):
    # type: (str, str, OrderedDict) -> None
    """Updates a registry in the csv file.

    Updates existing registries in the csv file using the name as the
    identifier.

    Parameters:
        filepath (str): The location of the csv file to add data.
        name (str): Identifier for the editing registry.
        new_data (OrderedDict): Ordered dictionary with the data.

    Returns:
        None.
    """

    # Read file
    try:
        with open(filepath, "r") as f:

            # Create a DictReader object
            reader = csv.DictReader(f)

            # Store header and list in separate variables
            headers = reader.fieldnames
            data_rows = list(reader)

        with open(filepath, 'wb') as f:

            # Create a DictWriter object
            writer = csv.DictWriter(f, fieldnames=headers)

            # Write header row
            writer.writeheader()

            # Upddate datarow with specified ID | name
            for row in data_rows:
                if row['File name'] == name:
                    row.update(new_data)

                    # Show updated record window
                    rs.MessageBox(
                        message="Succesfully updated model data:\n{}"
                        .format(row),
                        buttons=0,
                        title="Data updated"
                    )
                # Rewrite all rows
                writer.writerow(row)

    except IOError as error:
        print("Error reading or updating the CSV file", error)


# ----- DEBUGGING ROUTINE -----
# if __name__ == "__main__":

#     test_folder = os.path.dirname(
#         r"C:\Users\gdleraya\Box\Tooling Dev Center - Business Case\9. Projects\Guad North\03_Ciena\DM-GDN03-220050\CAM")

#     folder = rs.BrowseForFolder(
#         folder=test_folder
#     )
#     filepath = os.path.join(folder, "DATA.csv")

#     try:
#         print(get_row_index(filepath))
#     except:
#         print("An error ocurred")
