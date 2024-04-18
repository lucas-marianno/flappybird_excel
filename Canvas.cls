VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Canvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private skyHeight, groundHeight, leftEdge, rightEdge As Integer
Private rowHeight, columnWidth As Double
Private skyRng, groundRng, menuRng As range
Private menuInfo As String

Public Sub Init()
    skyHeight = 10
    groundHeight = skyHeight + 70
    leftEdge = 12
    rightEdge = leftEdge + 60
    rowHeight = 7
    columnWidth = 2.3
    
    menuInfo = "{F4} = Menu" & vbNewLine & _
                    "{ESPAÇO} = Flap"
    
    With Worksheets("game")
        Set skyRng = .range(.Cells(skyHeight, leftEdge), .Cells(groundHeight - 1, rightEdge))
        Set groundRng = .range(.Cells(groundHeight, leftEdge), .Cells(groundHeight, rightEdge))
        Set menuRng = .range(.Cells(1, leftEdge), .Cells(skyHeight - 1, rightEdge))
    End With

    Call SetUpSheet
    Call DrawCanvas
End Sub
Public Property Get Ground() As Integer
    Ground = groundHeight
End Property
Public Property Get Sky() As Integer
    Sky = skyHeight
End Property
Public Property Get EdgeR() As Integer
    EdgeR = rightEdge
End Property
Public Property Get EdgeL() As Integer
    EdgeL = leftEdge
End Property
Public Sub DrawCanvas()
    
    With skyRng
        .Interior.Color = 15773696
        .ClearContents
    End With
    
    With groundRng
        .Interior.Color = 5296274
        .ClearContents
    End With
    
End Sub
Public Sub Clear()
    With Worksheets("game")
        With .Cells
            .Clear
            .columnWidth = 8.43
            .rowHeight = 15
        End With
        
        .range("F20").Value = "Aperta 'F4' aí"
    End With
    
End Sub

Private Sub SetUpSheet()
    
    With Worksheets("game")
        With .Cells
            .ClearContents
            .Interior.Color = xlThemeColorLight1
            .columnWidth = columnWidth
            .rowHeight = rowHeight
        End With
        
        With menuRng
            .Merge
            .Value = menuInfo
            .Font.ThemeColor = xlThemeColorDark1
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        .range("A1").Select
    End With
End Sub
