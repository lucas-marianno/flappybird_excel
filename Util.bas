Attribute VB_Name = "Util"
Sub Sleep(milliseconds As Long)
    Dim endTime As Double
    
    ' Calculate the end time by adding the milliseconds to the current time
    endTime = Timer + (milliseconds / 1000) ' Convert milliseconds to seconds
    
    ' Loop until the current time reaches the end time
    Do While Timer < endTime
        ' Do nothing, just wait
        DoEvents ' Allows other events to process during the wait
    Loop
End Sub
