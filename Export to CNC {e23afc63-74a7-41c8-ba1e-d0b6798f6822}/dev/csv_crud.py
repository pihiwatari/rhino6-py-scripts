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
    * read_project_data - reads project information from the designated data file.
    * get_row_index - returns the last row index as reference for the next item to
    be added in the csv file.
    * add_data_row - Adds a new line of data to the csv file. Adds a single new row 
    to a csv file with the information of a RhinoGeometry object.
    * update_data_row - Edits an existing row in the csv file.
"""

import rhinoscriptsyntax as rs
from collections import OrderedDict
import csv
import os
import re


# This list can change in order or number of items present without notice, if the order
# and quantity of item doesn't match to the destination, please let me now:
# elias.rayas@flex.com

csv_headers = [
    "Row",
    "Flex Customer",
    "Site",
    "Building",
    "ID ATC",
    "Model or customer ID",
    "Product",
    "File name",
    "Requester",
    "Empty1",  # Start date
    "Empty2",  # Commit date
    "Material",
    "Designer",
    "Sizes",
    "Quantity",
    "Type",
    "Empty3",  # Status CNC machine
    "Empty4",  # Status assembly
    "Empty5",  # Delivered
    "Empty6",  # CAM date
    "Empty7",  # Programer
    "Empty8",  # Machining CNC date
    "Empty9",  # Machine
    "Empty10",  # End of CNC machining
    "Empty11",  # Machining time
    "File Path",  # Box drive link
    "Special details"
]


def generate_atc_id():
    # type: () -> str | None
    """Creates a new csv file.

    Generates a filename string using a regex for the current working
    Rhino document, following ATC naming conventions. Returns None if
    no name is input.
    """

    regex = re.compile("(DM-\D{1,3}\d{1,3}-\d{1,6})")
    filename = rs.DocumentName()
    if filename:
        match = re.search(regex, filename)
        if match:
            atc_id = match.group(0)
            return atc_id
    return


# Pre-filled columns, once the DATA file is created, the script will read
# from this entries/columns.
project_columns = {
    "Flex Customer": "Flex",
    "Site": "Guad North",
    "Building": "B18",
    "ID ATC": generate_atc_id(),  # Gets current filename as default value
    "Model or customer ID": "",
    "Product": "Product name",
    "Requester": "Requester name",
    "Designer": "Elias Rayas",
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
        print("DATA file found!")
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
    """ Creates an OrderedDict with predefined ordered key values.

    Returns a dictionary with predefined key values. All of the dict values
    are None by default.

    Parameters:
        None.

    Returns:
        data_dict (OrderedDict): dictionary with the predefined key values.
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

        # Open dir folder to extract box drive's link
        if key == "File Path":
            rs.BrowseForFolder(
                title="Create box drive link",
                message="Generate the folder's box drive link"
            )

        # Prompt for user input on the data.
        data = rs.StringBox(
            message="Please input information for: %s." % key,
            title="Setup project data",
            default_value=default_value
        )

        if data == None or data == 'None':
            raise ValueError("No data inputted! Cancelling sequence")

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
    except IOError as e:
        print("""
        Error in function: read_project_data.\n 
        This csv has no data to read.""")
        raise (e)


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

    # Destructure object to display model name, material and quantity once
    # finished writing the data row in the csv file.
    model_name = data_dict["File name"]
    model_material = data_dict["Material"]
    model_quantity = int(data_dict["Quantity"])

    try:
        with open(filepath, "ab") as f:

            writer = csv.DictWriter(f, fieldnames=csv_headers)
            writer.writerow(data_dict)
            rs.MessageBox(
                message="Succesfully saved model data:\n"
                "Filename: {}\n"
                "Material: {}\n"
                "Quantity: {}\n"
                .format(model_name, model_material, model_quantity),
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
