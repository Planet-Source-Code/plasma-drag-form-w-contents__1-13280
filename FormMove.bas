Attribute VB_Name = "FormMove"
'Drags Form While Showing Contents
'Pretty Cool
'www.plazmuh.com

'<<<< BEGIN CODE >>>>
'Private Sub Form_Load()
'FormMove.InitTPP
'End Sub
'
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then MouseDown
'End Sub
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then MouseMove Me
'End Sub
'<<<< END CODE >>>>

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Type POINTAPI
        X As Long
        Y As Long
End Type

Public LastPoint As POINTAPI

Public lngTPPY As Long
Public lngTPPX As Long
Sub MouseDown()
    Dim POINT As POINTAPI
    
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
End Sub
Sub MouseMove(ctlForm As Form)
    Dim lngX     As Long
    Dim lngY     As Long
    Dim POINT    As POINTAPI
       
    GetCursorPos POINT
    lngX& = (POINT.X - LastPoint.X) * lngTPPX&
    lngY& = (POINT.Y - LastPoint.Y) * lngTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    ctlForm.Move ctlForm.Left + lngX&, ctlForm.Top + lngY&
End Sub
Sub InitTPP()
    lngTPPX& = Screen.TwipsPerPixelX
    lngTPPY& = Screen.TwipsPerPixelY
End Sub
