Attribute VB_Name = "pbmod"
Global adNum
Global StopScript
Global Log
Global StopNow
Global openbrowser
Global TimerSet
Global id
Global AdCount
Global UserInfo
Global SkipClick
Global PingSent
Sub TimeOut(duration)
    Dim StartTime, X
    StartTime = Timer
    Do While Timer - StartTime < duration
        X = DoEvents()
    Loop
End Sub
