VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Plate As MSForms.Frame
Public WithEvents Face As MSForms.Label

Public ParentForm As Object
Public ActionName As String
Public Host As Object
Public IdleBack As Long
Public DownBack As Long

Private mArmed As Boolean


Private Sub Plate_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    Arm Button

End Sub


Private Sub Face_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    Arm Button

End Sub


Private Sub Plate_MouseUp( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    FinishClick Button, X, Y, 0, 0

End Sub


Private Sub Face_MouseUp( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    FinishClick Button, X, Y, Face.Left, Face.Top

End Sub


Private Sub Arm(ByVal Button As Integer)

    If Button <> 1 Then Exit Sub
    mArmed = True
    PressColor

End Sub


Private Sub FinishClick( _
    ByVal Button As Integer, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal originLeft As Single, _
    ByVal originTop As Single)

    Dim px As Single
    Dim py As Single
    Dim inside As Boolean

    RestoreColor
    If Button <> 1 Then
        mArmed = False
        Exit Sub
    End If
    If Not mArmed Then Exit Sub
    mArmed = False
    If Host Is Nothing Then Exit Sub

    px = originLeft + X
    py = originTop + Y
    inside = (px >= 0 And py >= 0 And px <= Host.Width And py <= Host.Height)
    If inside Then Fire

End Sub


Private Sub PressColor()

    If Host Is Nothing Then Exit Sub
    If DownBack = IdleBack Then Exit Sub
    Host.BackColor = DownBack

End Sub


Private Sub RestoreColor()

    If Host Is Nothing Then Exit Sub
    Host.BackColor = IdleBack

End Sub


Private Sub Fire()

    If Not ParentForm Is Nothing Then
        ParentForm.ButtonAction ActionName
    End If

End Sub
