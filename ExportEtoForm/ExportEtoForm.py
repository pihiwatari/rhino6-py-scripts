import Rhino
import scriptcontext
import rhinoscriptsyntax as rs
import System
import Rhino.UI
import Eto.Drawing as drawing
import Eto.Forms as forms

# * = Repeating data over rows,
# _ = Empty data non-editable cells
# - = Auto-filling data, non-editable cells
col_headers = [
    '-ID',
    '-Object name',
    '*Flex Customer',
    '*Site',
    '*Building',
    '*ID ATC',
    '*Model or customer ID',
    '*Product',
    'File name',
    '*Requester',
    '_Start date',  # Start date
    '_Commit date',  # Commit date
    'Material',
    '*Designer',
    '-Sizes',
    'Quantity',
    '*Type',
    '_Status CNC machine',  # Status CNC machine
    '_Status assembly',  # Status assembly
    '_Delivered',  # Delivered
    '_CAM date',  # CAM date
    '_Programer',  # Programer
    '_Machining CNC date',  # Machining CNC date
    '_Machine',  # Machine
    '_End of CNC machine',  # End of CNC machining
    '_Machining date',  # Machining time
    'File Path',  # Box drive link
    'Special details'
]


### Auxiliary Functions ###

def create_column(header, index, editable=True, visible=True):
    column = forms.GridColumn()
    column.HeaderText = header
    column.DataCell = forms.TextBoxCell(index)
    column.Editable = editable
    column.Visible = visible
    return column


def create_initial_gridview_columns(headers, gridview):
    for i, h in enumerate(headers):
        if h[0] == '_':
            column = create_column(h[1::], i, editable=False, visible=False)
        elif h[0] == '-':
            column = create_column(h[1::], i, editable=False)
        else:
            column = create_column(h, i)
        gridview.Columns.Add(column)

### Rhino Geometry Functions ###


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


class SampleGridView(forms.Dialog[bool]):

    # Dialog box Class initializer
    def __init__(self):

        self.Title = 'Export to CNC form'
        self.Width = 1024
        self.Height = 600
        self.Resizable = True

        # Store object
        self.Collection = forms.FilterCollection[list]([])

        # Groupbox settings
        self.m_groupbox = forms.GroupBox(Text='Project data')
        self.m_groupbox.Padding = drawing.Padding(5)

        group_layout = forms.DynamicLayout()
        group_layout.Spacing = drawing.Size(5, 5)

        # Aggregate groupbox controls
        g_label1 = forms.Label(Text='Flex customer')
        g_textbox1 = forms.TextBox(PlaceholderText='e.g. Arista', Width=300)

        g_label2 = forms.Label(Text='Site')
        g_textbox2 = forms.TextBox(Text='Guadalajara North', Width=300)

        g_label3 = forms.Label(Text='Building')
        g_textbox3 = forms.TextBox(PlaceholderText='e.g. B18', Width=300)

        g_label4 = forms.Label(Text='ID ATC')
        g_textbox4 = forms.TextBox(
            PlaceholderText='e.g. DM-GDN08-240001', Width=300)

        g_label5 = forms.Label(Text='Model or customer ID')
        g_textbox5 = forms.TextBox(
            PlaceholderText='e.g. ARI-235045-123T', Width=300)

        group_layout.AddRow(g_label1, g_textbox1, g_label2, g_textbox2)
        group_layout.AddRow(g_label3, g_textbox3, g_label4, g_textbox4)
        group_layout.AddRow(g_label5, g_textbox5)

        # Add the groupbox to the layout
        self.m_groupbox.Content = group_layout
        ### --------------------------------------------------- ###

        ### GridView settings ###

        self.m_gridview = forms.GridView()
        self.m_gridview.ShowHeader = True
        self.m_gridview.BackgroundColor = drawing.Color.FromGrayscale(250, 255)
        self.m_gridview.DataStore = self.Collection
        self.m_gridview.AllowMultipleSelection = True

        create_initial_gridview_columns(col_headers, self.m_gridview)

        # Create the default button
        self.AddGeometryButton = forms.Button(Text='Add new object')
        self.AddGeometryButton.Click += self.OnPushPickButton

        # Create a layout and add all the controls
        layout = forms.DynamicLayout()
        layout.Padding = drawing.Padding(10)
        layout.Spacing = drawing.Size(5, 5)
        layout.AddRow(self.m_groupbox)
        layout.AddRow(None)  # spacer
        layout.AddRow(self.AddGeometryButton)
        layout.AddRow(self.m_gridview)

        self.Content = layout

    ### Dialog methods ###

    def InitializeDataColletion(self):
        headers = []
        for col in range(len(self.m_gridview.Columns)):
            header = self.m_gridview.Columns[col].HeaderText
            headers.append(header)
        first_row = ["" for item in headers]
        first_row[0] = "1"
        return first_row

    def AddNewGeo(self, sender, e):

        try:
            # Create initial row
            objects, bb_dim = self.OnSelectNewObject()
            if self.m_gridview.DataStore.Count == 0:
                new_row = self.InitializeDataColletion()
            else:

                # Get last row data
                last_row = self.m_gridview.DataStore.Item[self.m_gridview.DataStore.Count-1]

                # Create an empty list to store new data
                new_row = []
                new_row.append(str(int(last_row[0])+1))

                # Copy last row project information
                for index in range(1, len(self.m_gridview.Columns)):
                    if col_headers[index][0] == '*':
                        new_row.append(str(last_row[index]))
                    else:
                        new_row.append('')

            # Both last_row and mew_row length must match
            self.m_gridview.DataStore.Add(new_row)
        except:
            print("An error has occurred!")
            return

    # Select Rhino Geometry
    def OnSelectNewObject(self):
        objects = rs.GetObjects(
            "Select objects new Object", preselect=True)  # Returns list
        if not objects:
            print("No objects selected, sequence cancelled!")
            return
        # Get bounding box dimensions for each object
        bb_dim = get_bb_dimensions(objects)
        print(objects)
        return (objects, bb_dim)

    # Hide dialog during select object command
    def OnPushPickButton(self, sender, e):
        Rhino.UI.EtoExtensions.PushPickButton(self, self.AddNewGeo)

    # Update project data on all rows


def Run():
    dialog = SampleGridView()
    rc = dialog.ShowModal(Rhino.UI.RhinoEtoApp.MainWindow)


if __name__ == '__main__':
    Run()
