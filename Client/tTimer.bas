Attribute VB_Name = "tTimer"
Option Explicit

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Sub MyTimer(hwnd As Long, msg As Long, idTimer As Long, dwTime As Long)
Static i As Integer

    i = i + 1
    Winsock.tIncrease
End Sub


