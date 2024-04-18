VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pipes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private positionY, pipeTop, pipeBottom  As Integer
Private gapSize, gapTop, gapBottom As Integer

Public Sub Init()
    
    positionY = myCanvas.EdgeR
    pipeTop = myCanvas.Sky
    pipeBottom = myCanvas.Ground - 1
    
    gapSize = 20 'ajust this to handle difficulty
    gapTop = 25 'randomize this to get different pipes
    gapBottom = gapTop + gapSize
End Sub

Public Sub DrawPipe()
    Dim topPipeRng, bottomPipeRng As range
    
    With Worksheets("game")
        Set topPipeRng = .range(.Cells(pipeTop, positionY), .Cells(gapTop, positionY))
        Set bottomPipeRng = .range(.Cells(gapBottom, positionY), .Cells(pipeBottom, positionY))
    End With
    
    topPipeRng.Interior.Color = 5287936
    bottomPipeRng.Interior.Color = 5287936
    
End Sub
Public Sub Move()
    positionY = positionY - 1
    
    CheckColitionBird
    CheckLeftEdge
    
End Sub

Private Sub CheckColitionBird()
    'handle colition logic
End Sub
Private Sub CheckLeftEdge()
    'handle check
    If positionY <= myCanvas.EdgeL Then
        UnDrawPipe
    End If
End Sub
Private Sub UnDrawPipe()
    'handle undraw
    positionY = myCanvas.EdgeR
End Sub
