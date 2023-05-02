import rhinoscriptsyntax as rs
import subprocess

__commandname__ = "GetBoundingBoxDimensions"

# RunCommand is the called when the user enters the command name in Rhino.
# The command name is defined by the filname minus "_cmd.py"


def copy2clip(txt):
    cmd = 'echo ' + txt.strip() + '|clip'
    return subprocess.check_call(cmd, shell=True)


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

    copy2clip(dimensions)
    print("Copied to clipboard: ", dimensions)
    return dimensions


def RunCommand(is_interactive):

    # Select object
    geo = rs.GetObjects(
        message="Select objects to calculcate dimensionss",
        select=True,
        preselect=True
    )

    if not geo:
        return 1

    get_bb_dimensions(geo)


# if __name__ == "__main__":
#     RunCommand(False)
