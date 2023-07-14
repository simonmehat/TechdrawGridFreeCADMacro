# Take the selected element on the techdraw view

global dvp
global SpaceBetweenGridAndDimension
global DictionaryCharactersSurrounded
DictionaryCharactersSurrounded = {'0': '&#9450;', '1': '&#9312;', '2': '&#9313;', '3': '&#9314;', '4': '&#9315;', '5': '&#9316;', '6': '&#9317;', '7': '&#9318;', '8': '&#9319;', '9': '&#9320;', '10': '&#9321;', '11': '&#9322;', '12': '&#9323;', '13': '&#9324;', '14': '&#9325;', '15': '&#9326;', '16': '&#9327;', '17': '&#9328;', '18': '&#9329;', '19': '&#9330;', '20': '&#9331;', '21': '&#12881;', '22': '&#12882;', '23': '&#12883;', '24': '&#12884;', '25': '&#12885;', '26': '&#12886;', '27': '&#12887;', '28': '&#12888;', '29': '&#12889;', '30': '&#12890;', '31': '&#12891;', '32': '&#12892;', '33': '&#12893;', '34': '&#12894;', '35': '&#12895;', '36': '&#12977;', '37': '&#12978;', '38': '&#12979;', '39': '&#12980;', '40': '&#12981;', '41': '&#12982;', '42': '&#12983;', '43': '&#12984;', '44': '&#12985;', '45': '&#12986;', '46': '&#12987;', '47': '&#12988;', '48': '&#12989;', '49': '&#12990;', '50': '&#12991;', 'A': '&#9398;', 'B': '&#9399;', 'C': '&#9400;', 'D': '&#9401;', 'E': '&#9402;', 'F': '&#9403;', 'G': '&#9404;', 'H': '&#9405;', 'I': '&#9406;', 'J': '&#9407;', 'K': '&#9408;', 'L': '&#9409;', 'M': '&#9410;', 'N': '&#9411;', 'O': '&#9412;', 'P': '&#9413;', 'Q': '&#9414;', 'R': '&#9415;', 'S': '&#9416;', 'T': '&#9417;', 'U': '&#9418;', 'V': '&#9419;', 'W': '&#9420;', 'X': '&#9421;', 'Y': '&#9422;', 'Z': '&#9423;'}
SpaceBetweenGridAndDimension = float(7.5)
dvp = App.ActiveDocument.ActiveObject
App.ActiveDocument.save()

print("---- Start Macro grid on techdraw ----")

def FindVertexPosition(VertexNumber):
     # Align
    VertexString = dvp.getVertexByIndex(VertexNumber).dumpToString()  # display many shape info
    VertexPositionStart = int(VertexString.find('3D :'))
    VertexPositionEnd = int(VertexString.find('\n', VertexPositionStart))
    Position3Value = VertexString[VertexPositionStart + 5: VertexPositionEnd]
    # cut the position in X and Y position
    FirstComa = Position3Value.find(",")
    SecondComa = Position3Value.find(",", FirstComa + 1)
    XposCalc = float(Position3Value[0:FirstComa])
    YposCalc = float(Position3Value[FirstComa + 1:SecondComa])
    return [XposCalc, YposCalc]


# search the left point
def OriginCorrection():
    print("---OrigineCorrection")
    Count = 0
    Xpos = 0
    Ypos = 0
    while True:
        try:
            XposCalc = FindVertexPosition(Count)[0]
            YposCalc = FindVertexPosition(Count)[1]

            if XposCalc <= Xpos and YposCalc <= Ypos:
                Xpos = XposCalc
                Ypos = YposCalc
            Count = Count + 1
            VertexString = ""
        except:
            break

    # Take 3D box origin placement
    OriginBox = App.ActiveDocument.getObjectsByLabel("origin")[0]
    Origin = str(OriginBox.Placement.Base)
    FirstComa = Origin.find(",")
    SecondComa = Origin.find(",", FirstComa + 1)
    OriginX = float(Origin[8:FirstComa - 1])
    OriginY = float(Origin[FirstComa + 2: SecondComa - 1])
    # corrects the offset between the grid origin (of 3D space) and the leftmost element of the sheet
    Xpos -= OriginX
    Ypos -= OriginY
    PositionList = [Xpos, Ypos]
    print("Position list : ", PositionList)
    return PositionList


def GetSpreadsheetGridDimension(SpreadsheetName, Columne, angle):  # transform grid position in a spreadsheet in a list
    print("---GetSpreadsheetGridDimension")
    Sheet = App.ActiveDocument.getObjectsByLabel(SpreadsheetName)[0]  # take the sheet were they are value of the grid
    ListOfGridGlobalPosition = [angle]
    NumCell = int(1)
    Cell = (Columne + "1")
    while str(Cell) in App.ActiveDocument.Grid.PropertiesList:
        Coord = Sheet.get(Cell)
        NumCell += 1
        ListOfGridGlobalPosition.append(Coord)
        Cell = (Columne + str(NumCell))

    Last_A_cell = int(NumCell - 1)
    print("ListOfGridGlobalPosition : ", ListOfGridGlobalPosition)
    return ListOfGridGlobalPosition


def DrawingBoundingBox(XY):  # define the length of the grid in function of the size of the view
    print("---DrawingBoundingBox")
    if XY == "X":
        Lenght = App.ActiveDocument.getObjectsByLabel("origin")[0].Y_grid_size
    if XY == "Y":
        Lenght = App.ActiveDocument.getObjectsByLabel("origin")[0].X_grid_size
    Lenght = float(str(Lenght)[:str(Lenght).find("m") - 1])
    print("Lenght : ", Lenght)
    return Lenght


def TechDrawGridLine(ListOfGridGlobalPosition, Xpos, Ypos, Lenght):  # create cosmetic line grid on TechDraw, with list
    def DefineStartEndPointOfCosmeticLine(GridValue, Xpos, Ypos, Lenght, Move):
        print("---DefineStartEndPointOfCosmeticLine")
        if ListOfGridGlobalPosition[0] == 90:
            Start = FreeCAD.Vector(GridValue + Xpos, 0 + Ypos + Move, 0)
            End = FreeCAD.Vector(GridValue + Xpos, Lenght + Ypos + Move, 0)
        elif ListOfGridGlobalPosition[0] == 0:
            Start = FreeCAD.Vector(0 + Xpos + Move, GridValue + Ypos, 0)
            End = FreeCAD.Vector(Lenght + Xpos + Move, GridValue + Ypos, 0)
        ListStartEnd = [Start, End]
        return ListStartEnd

    def DrawLineWithPoint(ListStartEnd):
        print("---DrawLineWithPoint")
        style = 4
        weight = 0.75
        pyGreen = (0.0, 0.0, 0.0, 1.0)
        Start = ListStartEnd[0]
        End = ListStartEnd[1]
        dvp.makeCosmeticVertex(Start)
        dvp.makeCosmeticLine(Start, End, style, weight, pyGreen)
        if i >= 1:
            Count = 0
            while True:
                try:
                    V = dvp.getVertexByIndex(Count)
                    # search the position of the vertex
                    Count = Count + 1
                except:
                    AddGridDimension(Count - 1)
                    break
        Count = 0
        while True:
            try:
                V = dvp.getVertexByIndex(Count)
                # search the position of the vertex
                Count = Count + 1
            except:
                if ListOfGridGlobalPosition[0]==0 :
                    AddGridAnnotation(Count-1, GetSpreadsheetGridDimension("Grid", "A", 0),i+1 )
                elif ListOfGridGlobalPosition[0]==90:
                    AddGridAnnotation(Count-1, GetSpreadsheetGridDimension("Grid", "C", 90),i+1 )
                break

    def DrawLine(ListStartEnd):
        print("---DrawLine")
        style = 4
        weight = 0.75
        pyGreen = (0.0, 0.0, 0.0, 1.0)
        # place cosmetic line
        Start = ListStartEnd[0]
        End = ListStartEnd[1]
        dvp.makeCosmeticLine(Start, End, style, weight, pyGreen)

    print("---TechDrawGridLine")

    for i in range(len(ListOfGridGlobalPosition) - 1):
        GridValue = ListOfGridGlobalPosition[i + 1]
        if Lenght < 9000:
            DrawLineWithPoint(DefineStartEndPointOfCosmeticLine(GridValue, Xpos, Ypos, Lenght, 0))
        elif Lenght >= 9000:
            LineNumber = 1
            while Lenght/LineNumber >= 9000:
                LineNumber += 1
            LineSize = Lenght / LineNumber
            print("Ligne : ", LineSize,"*", LineNumber)
            DrawLineWithPoint(DefineStartEndPointOfCosmeticLine(GridValue, Xpos, Ypos, LineSize, 0))
            for i in range(LineNumber-1):
                Move = (i + 1) * LineSize
                DrawLine(DefineStartEndPointOfCosmeticLine(GridValue, Xpos, Ypos, LineSize, Move))


def AddGridDimension(VertexNumber):
    print("---AddGridDimension")
    Dim = FreeCAD.ActiveDocument.addObject('TechDraw::DrawViewDimension', 'GridDimension')
    Dim.Type = "Distance"
    Dim.References2D = [(dvp, ('Vertex' + str(VertexNumber))), (dvp, ('Vertex' + str(VertexNumber - 1)))]

    SearchElementList = dvp.InList
    SearchElementList[:] = list(map(str, SearchElementList))
    App.ActiveDocument.getObjectsByLabel(dvp.InList[SearchElementList.index("<DrawPage object>")].Label)[0].addView(Dim)

    # Align
    XposCalc = FindVertexPosition(VertexNumber)[0]
    YposCalc = FindVertexPosition(VertexNumber)[1]

    XposCalc2 = FindVertexPosition(VertexNumber-1)[0]
    YposCalc2 = FindVertexPosition(VertexNumber-1)[1]

    if XposCalc == XposCalc2:
        Dim.X = XposCalc * dvp.Scale - SpaceBetweenGridAndDimension
        Dim.Y = (YposCalc + ((YposCalc2 - YposCalc) / 2)) * dvp.Scale

    elif YposCalc == YposCalc2:
        Dim.Y = YposCalc * dvp.Scale - SpaceBetweenGridAndDimension
        Dim.X = (XposCalc + ((XposCalc2 - XposCalc) / 2)) * dvp.Scale



def AddGridAnnotation(VertexNumber, ListOfAnnotatioName, GridLineNumber):
    print("---AddAnnotation")
    print(VertexNumber, ListOfAnnotatioName, GridLineNumber)
    Annotation = FreeCAD.ActiveDocument.addObject('TechDraw::DrawRichAnno','GridAnnotation')
    Annotation.AnnoParent=dvp
    SearchElement='<DrawPage object>'

    SearchElementList = dvp.InList
    SearchElementList[:] = list(map(str, SearchElementList))
    App.ActiveDocument.getObjectsByLabel(dvp.InList[SearchElementList.index("<DrawPage object>")].Label)[0].addView(Annotation)

    Annotation.AnnoText = str("""<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd"><html><head><meta name="qrichtext" content="1" /><style type="text/css">p, li { white-space: pre-wrap; }</style></head><body style=" font-family:'Ubuntu'; font-size:15pt; font-weight:400; font-style:normal;"><p style=" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">"""+str(DictionaryCharactersSurrounded[str(ListOfAnnotatioName[GridLineNumber])])+"""</p></body></html>""")
    Annotation.ShowFrame=False

    XposCalc = FindVertexPosition(VertexNumber)[0]
    YposCalc = FindVertexPosition(VertexNumber)[1]

    if ListOfAnnotatioName[0]==0:
        Annotation.X = XposCalc * dvp.Scale - SpaceBetweenGridAndDimension*2
        Annotation.Y = (YposCalc -1) * dvp.Scale
        print(str(ListOfAnnotatioName[GridLineNumber]), str(XposCalc * dvp.Scale - SpaceBetweenGridAndDimension*2),str((YposCalc -1) * dvp.Scale ))

    elif ListOfAnnotatioName[0]==90:
        Annotation.Y = YposCalc * dvp.Scale - SpaceBetweenGridAndDimension*2
        Annotation.X = (XposCalc  -1) * dvp.Scale
        print(str(ListOfAnnotatioName[GridLineNumber]), str((XposCalc  -1) * dvp.Scale),str(YposCalc * dvp.Scale - SpaceBetweenGridAndDimension*2 ))

PositionList = OriginCorrection()
Xpos = PositionList[0]
Ypos = PositionList[1]
TechDrawGridLine(GetSpreadsheetGridDimension("Grid", "B", 0), Xpos, Ypos, DrawingBoundingBox("Y"))
TechDrawGridLine(GetSpreadsheetGridDimension("Grid", "D", 90), Xpos, Ypos, DrawingBoundingBox("X"))


"""
# commentaire

2 while true à supprimer : List[-1] : logueur au dernier élement de l'autre liste

# Update the grid

"""
print("---- End of Macro grid on techdraw ----")