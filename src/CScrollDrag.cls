VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CScrollDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Track As MSForms.Label
Public WithEvents Thumb As MSForms.Label

Public ParentForm As Object
Public ListName As String

Private mDown As Boolean
Private mOffset As Single


Private Sub Track_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    If Button <> 1 Then Exit Sub
    If ParentForm Is Nothing Then Exit Sub

    ParentForm.JumpScrollThumb ListName, Y

End Sub


Private Sub Thumb_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    If Button <> 1 Then Exit Sub

    mDown = True
    mOffset = Y
    Thumb.BackColor = RGB(100, 102, 108)

End Sub


Private Sub Thumb_MouseMove( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    If Not mDown Then Exit Sub

    If Button <> 1 Then
        mDown = False
        Thumb.BackColor = RGB(132, 134, 140)
        Exit Sub
    End If

    If ParentForm Is Nothing Then Exit Sub
    ParentForm.MoveScrollThumb ListName, Y - mOffset

End Sub


Private Sub Thumb_MouseUp( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    mDown = False
    If Not Thumb Is Nothing Then
        Thumb.BackColor = RGB(132, 134, 140)
    End If

End Sub
