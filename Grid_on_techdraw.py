# define Variable which will be use in several function :
global dvp
global SpaceBetweenGridAndDimension
global DictionaryCharactersSurrounded
DictionaryCharactersSurrounded = {'0': '&#9450;', '1': '&#9312;', '2': '&#9313;', '3': '&#9314;', '4': '&#9315;',
                                  '5': '&#9316;', '6': '&#9317;', '7': '&#9318;', '8': '&#9319;', '9': '&#9320;',
                                  '10': '&#9321;', '11': '&#9322;', '12': '&#9323;', '13': '&#9324;', '14': '&#9325;',
                                  '15': '&#9326;', '16': '&#9327;', '17': '&#9328;', '18': '&#9329;', '19': '&#9330;',
                                  '20': '&#9331;', '21': '&#12881;', '22': '&#12882;', '23': '&#12883;',
                                  '24': '&#12884;', '25': '&#12885;', '26': '&#12886;', '27': '&#12887;',
                                  '28': '&#12888;', '29': '&#12889;', '30': '&#12890;', '31': '&#12891;',
                                  '32': '&#12892;', '33': '&#12893;', '34': '&#12894;', '35': '&#12895;',
                                  '36': '&#12977;', '37': '&#12978;', '38': '&#12979;', '39': '&#12980;',
                                  '40': '&#12981;', '41': '&#12982;', '42': '&#12983;', '43': '&#12984;',
                                  '44': '&#12985;', '45': '&#12986;', '46': '&#12987;', '47': '&#12988;',
                                  '48': '&#12989;', '49': '&#12990;', '50': '&#12991;', 'A': '&#9398;', 'B': '&#9399;',
                                  'C': '&#9400;', 'D': '&#9401;', 'E': '&#9402;', 'F': '&#9403;', 'G': '&#9404;',
                                  'H': '&#9405;', 'I': '&#9406;', 'J': '&#9407;', 'K': '&#9408;', 'L': '&#9409;',
                                  'M': '&#9410;', 'N': '&#9411;', 'O': '&#9412;', 'P': '&#9413;', 'Q': '&#9414;',
                                  'R': '&#9415;', 'S': '&#9416;', 'T': '&#9417;', 'U': '&#9418;', 'V': '&#9419;',
                                  'W': '&#9420;', 'X': '&#9421;', 'Y': '&#9422;',
                                  'Z': '&#9423;'}  # A list with all surrounded character codes
# Set a distance between the end of the grid and the text of dimension
SpaceBetweenGridAndDimension = float(7.5)
# Set in dvp (drawviewpart) the selected element (the view in techdraw were we will add the grid)
dvp = App.ActiveDocument.ActiveObject
App.ActiveDocument.save()  # save the document

print("---- Start Macro grid on techdraw ----")


def FindVertexPosition(
        VertexNumber):  # this function find the position of vertex (a point on techdraw view) with the number of the vertex
    VertexString = dvp.getVertexByIndex(
        VertexNumber).dumpToString()  # stores information about the searched vertex in a variable, print(VertexString) for more informations
    # find in the VextexString, the start of position information
    VertexPositionStart = int(VertexString.find('3D :'))
    VertexPositionEnd = int(
        VertexString.find('\n', VertexPositionStart))  # find in the VextexString, the end of position information
    Position3Value = VertexString[
        VertexPositionStart + 5: VertexPositionEnd]  # put the position of the vertex as str in variable Position3Value print(Position3Value) for more informations

    # Next step is to extract X and Y position of the str VertexString
    # find in the VertexString the separator between X position and Y position
    FirstComa = Position3Value.find(",")
    SecondComa = Position3Value.find(",",
                                     FirstComa + 1)  # find in the VertexString the separator between Y position and Z position
    # Store the extract  X position in Variable
    XposCalc = float(Position3Value[0:FirstComa])
    # Store the extract  Y position in Variable
    YposCalc = float(Position3Value[FirstComa + 1:SecondComa])
    return [XposCalc, YposCalc]


def OriginCorrection():  # This function allows placing the grid on the 0 of the 3D view. It search the vertex of Origine box projected on techdraw view
    print("---OrigineCorrection")
    Count = 0
    Xpos = 0
    Ypos = 0
    while True:  # This part use a while True and a Try / except ! It must be modify
        # each loop, it tries to take the position of the vertex. If the vertex doesn't exist (or all the vertex are checked), it goes on the except.
        try:
            # try to take the position of the vertex
            XposCalc = FindVertexPosition(Count)[0]
            YposCalc = FindVertexPosition(Count)[1]

            # if the vertex is at the bottom left compared with the previous one
            if XposCalc <= Xpos and YposCalc <= Ypos:
                # save the position in other variable (position of Origine box in the techdraw view)
                Xpos = XposCalc
                Ypos = YposCalc
            Count = Count + 1
        except:
            break

    # Take 3D box origin placement
    OriginBox = App.ActiveDocument.getObjectsByLabel(
        "origin")[0]  # find the origin box in 3 view
    # recover the positon of origin view as a string of X,Y,Z value
    Origin = str(OriginBox.Placement.Base)
    # find in the VertexString the separator between X position and Y position
    FirstComa = Origin.find(",")
    # find in the VertexString the separator between Y position and Z position
    SecondComa = Origin.find(",", FirstComa + 1)
    # Store the extract  X position in Variable
    OriginX = float(Origin[8:FirstComa - 1])
    # Store the extract  Y position in Variable
    OriginY = float(Origin[FirstComa + 2: SecondComa - 1])

    # Corrects the offset between the grid origin (of 3D space) and the leftmost element of the sheet
    Xpos -= OriginX
    Ypos -= OriginY
    # here Xpos and Ypos will contain the position of 3D 0 on the techdraw View
    PositionList = [Xpos, Ypos]
    print("Position list : ", PositionList)
    return PositionList


def GetSpreadsheetGridDimension(SpreadsheetName, Columne,
                                angle):  # This function read a column of the grid spreadsheet and put the value in a list
    print("---GetSpreadsheetGridDimension")
    # take the sheet were they are value of the grid
    Sheet = App.ActiveDocument.getObjectsByLabel(SpreadsheetName)[0]
    # add the angle of the  column (For horizontal line of line is 0, for vertical is 90)
    ListOfGridGlobalPosition = [angle]
    NumCell = int(1)
    # use to create the number of cell (for example A1 or C4)
    Cell = (Columne + "1")
    # we go through the cells while they're still full
    while str(Cell) in App.ActiveDocument.Grid.PropertiesList:
        # recover the value set in the cell (correspond to the coordonate of the grid)
        Coord = Sheet.get(Cell)
        NumCell += 1
        ListOfGridGlobalPosition.append(Coord)  # add this value to the list
        Cell = (Columne + str(NumCell))  # go to the next cell

    print("ListOfGridGlobalPosition : ", ListOfGridGlobalPosition)
    return ListOfGridGlobalPosition


# define the length of the grid in function the parameter of origin box
def DrawingBoundingBox(XY):
    print("---DrawingBoundingBox")
    if XY == "X":  # if we want to recover the X length of the grid
        Length = App.ActiveDocument.getObjectsByLabel(
            "origin")[0].X_grid_size  # set Length on the origin X_grid_size
    if XY == "Y":  # if we want to recover the Y Length of the grid
        Length = App.ActiveDocument.getObjectsByLabel("origin")[0].Y_grid_size
    # remove the unit of measurement of the Length
    Length = float(str(Length)[:str(Length).find("m") - 1])
    print("Length : ", Length)
    return Length

# define the gap between the start point of the grid and the O of the grid


def GridGap(XY):
    print("---DrawingBoundingBox")
    if XY == "X":  # if we want to recover the X length of the grid
        Gap = App.ActiveDocument.getObjectsByLabel(
            "origin")[0].X_grid_gap  # set Length on the origin X_grid_size
    if XY == "Y":  # if we want to recover the Y Length of the grid
        Gap = App.ActiveDocument.getObjectsByLabel("origin")[0].Y_grid_gap
    # remove the unit of measurement of the Length
    Gap = float(str(Gap)[:str(Gap).find("m") - 1])
    print("Gap : ", Gap)
    return Gap


# This function create cosmetic line grid on TechDraw, with list of line position, the position of 0 projected on view, the Length of the grid
def TechDrawGridLine(ListOfGridGlobalPosition, Xpos, Ypos, Length, Gap):
    # This function
    # find the Start and End point pf the grid line (because to create cosmetic line, we must have 2 point)
    Length += Gap

    def DefineStartEndPointOfCosmeticLine(GridValue, Xpos, Ypos, Length, Move, Gap):
        print("---DefineStartEndPointOfCosmeticLine")
        if ListOfGridGlobalPosition[0] == 90:  # if the line is vertical
            Start = FreeCAD.Vector(GridValue + Xpos, 0 + Ypos + Move - Gap, 0)
            End = FreeCAD.Vector(
                GridValue + Xpos, Length + Ypos + Move - Gap, 0)
        elif ListOfGridGlobalPosition[0] == 0:  # if the line is horizontal
            Start = FreeCAD.Vector(0 + Xpos + Move - Gap, GridValue + Ypos, 0)
            End = FreeCAD.Vector(Length + Xpos + Move -
                                 Gap, GridValue + Ypos, 0)
        ListStartEnd = [Start, End]
        return ListStartEnd

    # this function draw one grid line and a point (the point will be use by the dimension)
    def DrawLineWithPoint(ListStartEnd):
        print("---DrawLineWithPoint")
        style = 4  # style 4 is Dash Dot line
        weight = 0.18  # the thickness of the line
        # the color of the line (here it is blue)
        pyGreen = (0.0, 0.0, 1.0, 0.0)
        Start = ListStartEnd[0]
        End = ListStartEnd[1]
        dvp.makeCosmeticVertex(Start)  # Draw a point at start position
        # Draw a line between start and end position
        dvp.makeCosmeticLine(Start, End, style, weight, pyGreen)
        if i >= 1:  # if it is not the first line of the grid
            Count = 0
            # we search th position of the last vertex (correspond to the cosmeticvertex create just above)
            while True:
                try:
                    V = dvp.getVertexByIndex(Count)
                    # search the position of the vertex
                    Count = Count + 1
                except:
                    # add dimension between this point and the point before
                    AddGridDimension(Count - 1)
                    break
        Count = 0
        # we add annotation add the start of the grid line
        while True:
            try:
                V = dvp.getVertexByIndex(Count)
                # search the position of the vertex (correspond to the cosmeticvertex create just above)
                Count = Count + 1
            except:
                if ListOfGridGlobalPosition[0] == 0:  # if the line is vertical
                    AddGridAnnotation(
                        Count - 1, GetSpreadsheetGridDimension("Grid", "A", 0), i + 1)  # take the name of annotation  dans put it on the drawing
                elif ListOfGridGlobalPosition[0] == 90:  # if it is horizontal
                    AddGridAnnotation(
                        Count - 1, GetSpreadsheetGridDimension("Grid", "C", 90), i + 1)
                break

    # draw line without point at the start, very similar to the start of function above
    def DrawLine(ListStartEnd):
        print("---DrawLine")
        style = 4
        weight = 0.18
        pyGreen = (0.0, 0.0, 1.0, 0.0)
        # place cosmetic line
        Start = ListStartEnd[0]
        End = ListStartEnd[1]
        dvp.makeCosmeticLine(Start, End, style, weight, pyGreen)

    print("---TechDrawGridLine")

    for i in range(len(ListOfGridGlobalPosition) - 1):  # scroll through the list positions
        GridValue = ListOfGridGlobalPosition[i + 1]
        if Length < 8000:  #  if the Length of grid line is under 8000
            DrawLineWithPoint(DefineStartEndPointOfCosmeticLine(
                GridValue, Xpos, Ypos, Length, 0, Gap))
        #  if the Length of the grid line is more than 8000 (this part exists because FreeCAD can't draw a line more than 10 000 mm Length ). Also, we will draw several line
        elif Length >= 8000:
            LineNumber = 1
            while Length / LineNumber >= 8000:  # try to find the good Length of the line, divisible by an integer
                LineNumber += 1
            LineSize = Length / LineNumber  # calcul the Length of little part of line
            print("Ligne : ", LineSize, "*", LineNumber)
            DrawLineWithPoint(DefineStartEndPointOfCosmeticLine(
                GridValue, Xpos, Ypos, LineSize, 0, Gap))  # draw the first line with a point (for the dimension)
            for i in range(LineNumber - 1):  # draw the followed line without point
                Move = (i + 1) * LineSize  # move the start of the line
                DrawLine(DefineStartEndPointOfCosmeticLine(
                    GridValue, Xpos, Ypos, LineSize, Move, Gap))


# Add dimension between the actual vertex and the last vertex
def AddGridDimension(VertexNumber):
    print("---AddGridDimension")
    Dim = FreeCAD.ActiveDocument.addObject(
        'TechDraw::DrawViewDimension', 'GridDimension')  # Add dimension object and change its name
    Dim.Type = "Distance"  # modify type of the dimension
    Dim.References2D = [(dvp, ('Vertex' + str(VertexNumber))),
                        (dvp, ('Vertex' + str(VertexNumber - 1)))]  # select the start and end point of the dimension

    SearchElementList = dvp.InList
    SearchElementList[:] = list(map(str, SearchElementList))
    App.ActiveDocument.getObjectsByLabel(
        dvp.InList[SearchElementList.index("<DrawPage object>")].Label)[0].addView(Dim)  # add the dimension to the techdraw view

    # Align the text of the dimension
    # recover the position of the vertex connected to the dimension
    XposCalc = FindVertexPosition(VertexNumber)[0]
    YposCalc = FindVertexPosition(VertexNumber)[1]

    XposCalc2 = FindVertexPosition(VertexNumber - 1)[0]
    YposCalc2 = FindVertexPosition(VertexNumber - 1)[1]

    if XposCalc == XposCalc2:  # if it is between 2 vertical line
        # set the position of dimension, take into account the scale
        Dim.X = XposCalc * dvp.Scale - SpaceBetweenGridAndDimension
        Dim.Y = (YposCalc + ((YposCalc2 - YposCalc) / 2)) * dvp.Scale

    elif YposCalc == YposCalc2:  # if it is between 2 horizontal line
        # set the position of dimension, take into account the scale
        Dim.Y = YposCalc * dvp.Scale - SpaceBetweenGridAndDimension
        Dim.X = (XposCalc + ((XposCalc2 - XposCalc) / 2)) * dvp.Scale


# this function add annotation of the view (the name of the line is for example A, or 1)
def AddGridAnnotation(VertexNumber, ListOfAnnotationName, GridLineNumber):
    print("---AddAnnotation")
    print(VertexNumber, ListOfAnnotationName, GridLineNumber)
    Annotation = FreeCAD.ActiveDocument.addObject(
        'TechDraw::DrawRichAnno', 'GridAnnotation')  # Add annotation element
    Annotation.AnnoParent = dvp  # attach it on the techdraw view

    SearchElementList = dvp.InList
    SearchElementList[:] = list(map(str, SearchElementList))
    App.ActiveDocument.getObjectsByLabel(dvp.InList[SearchElementList.index("<DrawPage object>")].Label)[0].addView(
        Annotation)  # add annotation to the page where there is the techdraw view

    Annotation.AnnoText = str(
        """<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd"><html><head><meta name="qrichtext" content="1" /><style type="text/css">p, li { white-space: pre-wrap; }</style></head><body style=" font-family:'Ubuntu'; font-size:15pt; font-weight:400; font-style:normal;"><p style=" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">""" + str(
            DictionaryCharactersSurrounded[str(ListOfAnnotationName[GridLineNumber])]) + """</p></body></html>""")  # Add text in HTML in the annotation. Use a dictionary  to place surrounded letter
    Annotation.ShowFrame = False

    # find the position of the vertex
    XposCalc = FindVertexPosition(VertexNumber)[0]
    YposCalc = FindVertexPosition(VertexNumber)[1]

    if ListOfAnnotationName[0] == 0:  # if grid line is horizontal
        Annotation.X = XposCalc * dvp.Scale - \
            SpaceBetweenGridAndDimension * 2  # place annotation
        Annotation.Y = (YposCalc - 1) * dvp.Scale
        print(str(ListOfAnnotationName[GridLineNumber]), str(XposCalc * dvp.Scale - SpaceBetweenGridAndDimension * 2),
              str((YposCalc - 1) * dvp.Scale))

    elif ListOfAnnotationName[0] == 90:  # if grid line is vertical
        Annotation.Y = YposCalc * dvp.Scale - \
            SpaceBetweenGridAndDimension * 2  # place annotation
        Annotation.X = (XposCalc - 1) * dvp.Scale
        print(str(ListOfAnnotationName[GridLineNumber]), str((XposCalc - 1) * dvp.Scale),
              str(YposCalc * dvp.Scale - SpaceBetweenGridAndDimension * 2))


PositionList = OriginCorrection()
Xpos = PositionList[0]
Ypos = PositionList[1]
TechDrawGridLine(GetSpreadsheetGridDimension("Grid", "B", 0),
                 Xpos, Ypos, DrawingBoundingBox("X"), GridGap("X"))
TechDrawGridLine(GetSpreadsheetGridDimension("Grid", "D", 90),
                 Xpos, Ypos, DrawingBoundingBox("Y"), GridGap("Y"))

print("---- End of Macro grid on techdraw ----")
