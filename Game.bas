Attribute VB_Name = "game"
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public myBird As New bird
Public myCanvas As New Canvas
Public myPipes As New Pipes

Sub Init()
    
    myCanvas.Init
    myBird.Init
    myPipes.Init

    Application.OnKey " ", "GameLoop"
End Sub
Sub GameMenu()
    form_GameMenu.Show
End Sub
Sub GameLoop()
    
    Do While myBird.IsAlive
    
        ' Other game logic goes here
        HandleInput
        UpdateGameLogic
        RenderGame
        
        ' Pause for a short time to control loop speed
        Sleep 100
    Loop
    
    If Not myBird.IsAlive Then
        MsgBox "ded"
        GameMenu
    End If
End Sub
Sub HandleInput()
    ' Check if spacebar (ASCII 32) is pressed
    If GetAsyncKeyState(32) Then
        myBird.Jump
    End If
End Sub
Sub UpdateGameLogic()
    ' Add code here to update game logic (e.g., handle user input, update object positions)
    myBird.Fall
    myPipes.Move
    
End Sub
Sub RenderGame()
    ' Add code here to render the game (e.g., draw objects on the worksheet)
    myCanvas.DrawCanvas
    myBird.DrawBird
    myPipes.DrawPipe
    
End Sub

Sub EndGame()
    Application.OnKey " ", ""
    myCanvas.Clear
End Sub

