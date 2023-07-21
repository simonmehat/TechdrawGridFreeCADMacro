# Create_origin

def CreateOriginBox():  # Create the orgin box
    box = App.ActiveDocument.addObject("Part::Box", "origin")  # create a box
    box.Length = 10  # set the dimension
    box.Width = 10
    box.Height = 10
    box.Placement = FreeCAD.Placement(
        FreeCAD.Vector(-100, -100, -200), FreeCAD.Rotation(0, 0, 0))  # place the box
    box.recompute()
    box = App.ActiveDocument.getObjectsByLabel("origin")[0]
    # add length of the grid property
    box.addProperty("App::PropertyDistance", "X_grid_size")
    box.addProperty("App::PropertyDistance", "Y_grid_size")
    box.X_grid_size = 2000  # set default value
    box.Y_grid_size = 2000


def CreateSpreadsheet():  # Create a spreadsheet
    sheet = App.ActiveDocument.addObject("Spreadsheet::Sheet", "Grid")
    sheet.set('B1', '0')  # add default value
    sheet.set('D1', '0')
    sheet.set('B2', '100')
    sheet.set('B3', '200')
    sheet.set('D2', '100')
    sheet.set('D3', '200')

    sheet.set('A1', 'A')
    sheet.set('A2', 'B')
    sheet.set('A3', 'C')
    sheet.set('C1', '1')
    sheet.set('C2', '2')
    sheet.set('C3', '3')

    sheet.recompute()


CreateOriginBox()
CreateSpreadsheet()
