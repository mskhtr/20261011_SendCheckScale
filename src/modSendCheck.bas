Attribute VB_Name = "modSendCheck"
Option Explicit

Public SendCheckScale As Single
Public KeptPeople As Collection
Public KeptFiles As Collection


Private Type SendCheckRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SendCheckMonitorInfo
    cbSize As Long
    rcMonitor As SendCheckRect
    rcWork As SendCheckRect
    dwFlags As Long
End Type

Private Declare PtrSafe Function SendCheckMonitorFromWindow Lib "user32" Alias "MonitorFromWindow" ( _
    ByVal hwnd As LongPtr, _
    ByVal dwFlags As Long) As LongPtr

Private Declare PtrSafe Function SendCheckGetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" ( _
    ByVal hMonitor As LongPtr, _
    ByRef lpmi As SendCheckMonitorInfo) As Long

Private Declare PtrSafe Function SendCheckGetDpiForWindow Lib "user32" Alias "GetDpiForWindow" ( _
    ByVal hwnd As LongPtr) As Long

Public Function SendCheckUiFits(ByVal uiScale As Single) As Boolean

    Dim hwnd As LongPtr
    Dim monitor As LongPtr
    Dim info As SendCheckMonitorInfo
    Dim dpi As Long
    Dim widthPt As Single
    Dim heightPt As Single

    If uiScale <= 0 Then uiScale = 1
    hwnd = CLngPtr(Application.Hwnd)
    monitor = SendCheckMonitorFromWindow(hwnd, 2)
    If monitor = 0 Then Exit Function
    info.cbSize = 40
    If SendCheckGetMonitorInfo(monitor, info) = 0 Then Exit Function
    dpi = SendCheckGetDpiForWindow(hwnd)
    If dpi <= 0 Then dpi = 96
    widthPt = (info.rcWork.Right - info.rcWork.Left) * 72 / dpi
    heightPt = (info.rcWork.Bottom - info.rcWork.Top) * 72 / dpi
    SendCheckUiFits = (widthPt >= 800 * uiScale) And (heightPt >= 460 * uiScale)

End Function


Public Sub KeepSendCheckState( _
    ByVal people As Collection, _
    ByVal files As Collection)

    Set KeptPeople = people
    Set KeptFiles = files

End Sub


Public Sub ReopenSendCheckScaled()

    送信前チェックを開く

End Sub

Public Sub 送信前チェックを開く()

    Dim uiScale As Single

    uiScale = SendCheckScale
    If uiScale <= 0 Then uiScale = 1

    If Not SendCheckUiFits(1) Then
        MsgBox _
            ChrW(&H3053) & ChrW(&H306E) & ChrW(&H753B) & ChrW(&H9762) & ChrW(&H306E) & _
            ChrW(&H89E3) & ChrW(&H50CF) & ChrW(&H5EA6) & ChrW(&H3067) & ChrW(&H306F) & _
            ChrW(&H3001) & ChrW(&H9001) & ChrW(&H4FE1) & ChrW(&H524D) & ChrW(&H30C1) & _
            ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H3092) & ChrW(&H8868) & _
            ChrW(&H793A) & ChrW(&H3067) & ChrW(&H304D) & ChrW(&H307E) & ChrW(&H305B) & _
            ChrW(&H3093) & ChrW(&H3002) & vbCrLf & _
            ChrW(&H5FC5) & ChrW(&H8981) & ChrW(&H306A) & ChrW(&H5927) & ChrW(&H304D) & _
            ChrW(&H3055) & ChrW(&H306F) & " 800 x 460 " & _
            ChrW(&H30DD) & ChrW(&H30A4) & ChrW(&H30F3) & ChrW(&H30C8) & ChrW(&H3067) & _
            ChrW(&H3059) & ChrW(&H3002), _
            vbExclamation, _
            ChrW(&H9001) & ChrW(&H4FE1) & ChrW(&H524D) & ChrW(&H30C1) & ChrW(&H30A7) & _
            ChrW(&H30C3) & ChrW(&H30AF)
        Exit Sub
    End If

    If Not SendCheckUiFits(uiScale) Then uiScale = 1
    SendCheckScale = uiScale

    On Error Resume Next
    Unload frmSendCheck
    On Error GoTo 0
    frmSendCheck.Show

End Sub