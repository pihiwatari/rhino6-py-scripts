import rhinoscriptsyntax as rs
import os
import re

import csv_crud as csvc

__commandname__ = "ExportToCNC"

# RunCommand is the called when the user enters the command name in Rhino.
# The command name is defined by the filname minus "_cmd.py"


def get_savefolder_name(save_location):
    # type: (str) -> str
    """
    Prompts user to select a folder to store exported 3D files,
    then it generates a project name based on folder path and returns
    both folder path and project name.

    Parameters:
        save_location (str): path to store exported 3D files and csv data.

    Returns:
        project_name (str): project name obtained from the save location
        path.
    """

    try:
        # Get project name from folder
        project_name = re.search(
            "DM-.*-.{6}", save_location)  # e.g. DM-GDN03-220001

        if not project_name:
            raise ValueError("Invalid project name")

        return project_name.group()

    except ValueError:
        rs.MessageBox(
            message="""
            Folder name must contain the following schema in its path:\n
            DM-XXX00-000000. Please try again.
            """,
            buttons=0,
            title="Invalid folder name"
        )

        return None


def create_filename(project_name, geo_guids):
    # type: (str, list[str]) -> str
    """Create filename using folder's name and RhinoGeometry name.

    In case the object has no name, request user input to add name to
    object and return the filename. All the selected objects will have
    the same name.

    Params:
        project_name (str): project name obtained from the save location
        path.
        geo_guids (list[str]): a list with the guid of rhino objects.
    """

    # Get rhinogeometry's name
    geo_name = rs.ObjectName(geo_guids[0])

    # Check for input name is not None
    if geo_name is None:
        geo_name = rs.StringBox(
            message="Input name for this object",
            title="Assign name to object"
        )
        for i in geo_guids:
            rs.ObjectName(i, geo_name)

    # Create name using both rhinogeometry name and project name
    filename = "{}_{}.stp".format(project_name, geo_name)

    return filename


def export_to_cnc(geo_name, folder_location):
    # type: (str, str) -> None
    """Custom export function with predefined format.

    Exports one object or several objects in one file, it relies on
    the rhino _Export command. This is only useful if exporting to a 
    folder with a DM-XXX00-00000 format in its path.

    Params:
        geo_name (str): string with the name of the object to export.
        folder_location (str): folder path to export object files. 

    Returns:
        None
    """

    # Create export filepath to use in export command
    filepath = os.path.join(folder_location, geo_name)

    # Export object to folder
    try:
        rs.Command('! _-Export "{}" _Enter'.format(filepath), echo=False)

    except:
        rs.MessageBox("Export Failed")
        return 1


def get_bb_dimensions(guids):
    # type: (list[str]) -> str
    """Returns objects' boundig box dimensions in WCS (x,y,z).

        Parameters:
            guid (list[str]): the guid of a rhino object.

        Returns:
            dimensions (str): the object's dimensions as a string in 
            "x * y * z" format.
    """

    # Create bounding box points
    bb_points = rs.BoundingBox(guids)

    # Calculate distance for each dimension
    x = rs.Distance(bb_points[0], bb_points[1])
    y = rs.Distance(bb_points[0], bb_points[3])
    z = rs.Distance(bb_points[0], bb_points[4])

    # Format returning string
    dimensions = "{:.2f} * {:.2f} * {:.2f}".format(x, y, z)

    print(dimensions)
    return dimensions


def RunCommand(is_interactive):

    ################################################################
    # ---MAIN ROUTINE EXECUTION--- #
    ################################################################
    print("Exporting...")

    # Select project's folder location
    save_location = rs.BrowseForFolder(
        title="Select folder to export"
    )

    if not save_location or save_location == None:
        print("No folder selected. Cancelling sequence...")
        return

    # Constants
    CSV_NAME = 'DATA.csv'
    PROJECT_ID = get_savefolder_name(save_location)

    if not PROJECT_ID:
        return 1
    ##################################
    # Create project CSV
    ##################################

    # Define csv file location
    csv_location = os.path.join(save_location, CSV_NAME)

    # Check for current DATA file existence, if it does not exist
    # then create a new one.
    if not os.path.exists(csv_location):
        csvc.create_new_csv(csv_location)

    # Create ordered data dictionary with the headers as 1st row
    data_dict = csvc.create_data_dict()

    # Read csv data, if the counter is > 1 it means there is information
    # in the csv, otherwise run the project data setup function.
    if csvc.get_row_index(csv_location) >= 1:

        # Read csv file data
        data_dict = csvc.read_project_data(csv_location)

    else:
        # Define project information for the CSV file, this information
        # is used on all rows of the document.
        csvc.setup_project_data(data_dict)

    # Declaring variables to get part data
    dimensions = None
    filename = None
    material = None
    quantity = None
    product_type = None

    # Assign product type
    product_type = rs.ComboListBox(
        items=["Pilot", "Production"],
        message="Select project type"
    )

    # Loop to get part all parts information, to cancel press ESC key.
    while True:

        print("Creating object data...")

        # Select object
        rhino_object = rs.GetObjects(
            message="Select object to export",
            select=True
        )

        if not rhino_object or rhino_object == None:
            return

        ##############################
        # This section of the script gets the object's data,
        # and adds it to the dictionary.
        ##############################
        # Get models bounding box
        dimensions = get_bb_dimensions(rhino_object)

        # Define filename
        filename = create_filename(PROJECT_ID, rhino_object)

        # Define material
        material = rs.StringBox(
            message="Add input for material",
            # default_value="SolidPro"
        )

        # Define machining quantities
        quantity = rs.RealBox(
            message="How many parts based of this geometry are needed?",
            default_number=1,
            minimum=1,
            maximum=100
        )

        ################################
        # This section of the script adds items to the dictionary, to be sent to
        # the csv file.
        ################################

        # Add dimensions to dictionary
        data_dict["Sizes"] = dimensions

        # Add project name to dictionary
        data_dict["File name"] = filename

        # Add Material to dictionary
        data_dict["Material"] = material

        # Add Quantity to  dictionary
        data_dict["Quantity"] = quantity

        # Add product type to item
        data_dict["Type"] = product_type

        ################################
        # Read and update csv items, or create new rows.
        ################################

        # Read CSV file to get last row index.
        row_index = csvc.get_row_index(csv_location, filename)

        # Update record if it exist in the csv, otherwise create a new
        # record with the data
        if row_index == -1:
            csvc.update_data_row(csv_location, filename, data_dict)

        # If this is a new item, then add a new row to the csv file.
        else:
            # Add row index to dictionary
            data_dict["Row"] = row_index

            # Write data to csv file.
            csvc.add_data_row(csv_location, data_dict)

        # Export model to save location
        export_to_cnc(filename, save_location)

        print("Successfully exported...")


if __name__ == "__main__":

    RunCommand(True)
