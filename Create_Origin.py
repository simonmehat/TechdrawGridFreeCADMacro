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

    # add grid_gap of the grid property
    box.addProperty("App::PropertyDistance", "X_grid_gap")
    box.addProperty("App::PropertyDistance", "Y_grid_gap")
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

    sheet.set('E1', 'CLT 0  -100')
    sheet.set('E2', 'CLT 1 downside floor  +2660')
    sheet.set('E3', 'CLT 1 upside floor  +2840')
    sheet.set('E4', 'CLT 2 downside floor  +5620')
    sheet.set('E5', 'CLT 2 upside floor  +5800')
    sheet.set('F1', '-100')
    sheet.set('F2', '2660')
    sheet.set('F3', '2840')
    sheet.set('F4', '5620')
    sheet.set('F5', '5800')

    sheet.setBackground('A1:A16384', (1.000000, 1.000000, 0.000000))
    sheet.setBackground('C1:C16384', (1.000000, 1.000000, 0.000000))
    sheet.setBackground('E1:E16384', (1.000000, 1.000000, 0.000000))
    sheet.recompute()


CreateOriginBox()
CreateSpreadsheet()
