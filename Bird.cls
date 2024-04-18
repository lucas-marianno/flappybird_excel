VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bird"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private bHeight As Integer
Private bIsAlive As Boolean
Private bMovement As Integer
Private bColumn As Integer
Private bRange As range

Public Sub Init()
    bHeight = myCanvas.Sky + 15
    bColumn = myCanvas.EdgeL + 5
    bMovement = 1
    bIsAlive = True
    
    Call DrawBird
End Sub
Public Property Get IsAlive() As Boolean
    IsAlive = bIsAlive
End Property
Public Property Get Height() As Integer
    Height = bHeight
End Property
Public Property Get Column() As Integer
    Column = bColumn
End Property


Public Sub BirdInfo()
    
    MsgBox "Value: " & bRange.Value & vbNewLine & _
            "Height: " & bHeight & vbNewLine & _
            "Alive: " & bIsAlive
    
End Sub
Public Sub Jump()
    
    bHeight = bHeight - 5
    bMovement = 1
    
End Sub

Private Sub OnGroundCheck()
    
    If bHeight >= myCanvas.Ground Then
        bHeight = myCanvas.Ground
        bIsAlive = False
    End If
    
End Sub
Public Sub DrawBird()
    
    If bHeight <= myCanvas.Sky Then
        bHeight = myCanvas.Sky
    End If
    
    Set bRange = Worksheets("game").Cells(bHeight, bColumn)
        
    With bRange
        .Interior.Color = 65535
        .Value = "''>'"
    End With
End Sub
Public Sub Fall()
    bHeight = bHeight + bMovement
    bMovement = bMovement * 2
    OnGroundCheck
End Sub


