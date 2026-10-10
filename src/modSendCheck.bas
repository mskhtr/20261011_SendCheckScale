Attribute VB_Name = "modSendCheck"
Option Explicit

Public SendCheckScale As Single
Public KeptPeople As Collection
Public KeptFiles As Collection

Private mReopenTimer As LongPtr


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

Private Declare PtrSafe Function SendCheckGetDpiForSystem Lib "user32" Alias "GetDpiForSystem" () As Long

Private Declare PtrSafe Function SendCheckSetTimer Lib "user32" Alias "SetTimer" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr) As LongPtr

Private Declare PtrSafe Function SendCheckKillTimer Lib "user32" Alias "KillTimer" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long

Public Function SendCheckUiFits(ByVal uiScale As Single) As Boolean

    Dim hwnd As LongPtr
    Dim monitor As LongPtr
    Dim info As SendCheckMonitorInfo
    Dim dpi As Long
    Dim widthPt As Single
    Dim heightPt As Single

    If uiScale <= 0 Then uiScale = 1
    hwnd = SendCheckHostHwnd()
    If hwnd = 0 Then
        monitor = SendCheckMonitorFromWindow(0, 1)
    Else
        monitor = SendCheckMonitorFromWindow(hwnd, 2)
    End If
    If monitor = 0 Then Exit Function
    info.cbSize = 40
    If SendCheckGetMonitorInfo(monitor, info) = 0 Then Exit Function
    dpi = 0
    If hwnd <> 0 Then dpi = SendCheckGetDpiForWindow(hwnd)
    If dpi <= 0 Then dpi = SendCheckGetDpiForSystem()
    If dpi <= 0 Then dpi = 96
    widthPt = (info.rcWork.Right - info.rcWork.Left) * 72 / dpi
    heightPt = (info.rcWork.Bottom - info.rcWork.Top) * 72 / dpi
    SendCheckUiFits = (widthPt >= 800 * uiScale) And (heightPt >= 460 * uiScale)

End Function



Private Function SendCheckMacroName() As String

    SendCheckMacroName = _
        ChrW(&H9001) & ChrW(&H4FE1) & ChrW(&H524D) & ChrW(&H30C1) & _
        ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H3092) & _
        ChrW(&H958B) & ChrW(&H304F)

End Function


Private Function SendCheckHostHwnd() As LongPtr

    Dim window As Object

    On Error Resume Next
    Set window = Application.ActiveWindow
    If Not window Is Nothing Then
        SendCheckHostHwnd = CLngPtr(window.HWND)
    End If
    If SendCheckHostHwnd <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set window = Application.ActiveExplorer
    If Not window Is Nothing Then
        SendCheckHostHwnd = CLngPtr(window.HWND)
    End If
    If SendCheckHostHwnd <> 0 Then
        Err.Clear
        Exit Function
    End If
    Set window = Application.ActiveInspector
    If Not window Is Nothing Then
        SendCheckHostHwnd = CLngPtr(window.HWND)
    End If
    Err.Clear

End Function


Public Sub ScheduleReopenSendCheck()

    If mReopenTimer <> 0 Then
        SendCheckKillTimer 0, mReopenTimer
        mReopenTimer = 0
    End If
    mReopenTimer = SendCheckSetTimer(0, 0, 700, AddressOf SendCheckReopenTimer)

End Sub


Public Sub SendCheckReopenTimer( _
    ByVal hwnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal nIDEvent As LongPtr, _
    ByVal dwTime As Long)

    SendCheckKillTimer 0, nIDEvent
    mReopenTimer = 0
    送信前チェックを開く

End Sub


Public Sub SendCheckEnsureButton()

    Dim exp As Object
    Dim bar As Object
    Dim ctrl As Object
    Dim btn As Object
    Dim caption As String

    On Error Resume Next
    Set exp = Application.ActiveExplorer
    If exp Is Nothing Then
        Err.Clear
        Exit Sub
    End If

    caption = SendCheckMacroName()
    Set bar = exp.CommandBars.Item("SendCheck")
    If bar Is Nothing Then
        Set bar = exp.CommandBars.Add("SendCheck", 1, False, False)
    End If
    If bar Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    bar.Visible = True

    For Each ctrl In bar.Controls
        If ctrl.Tag = "SendCheckScale" Then
            ctrl.Caption = caption
            ctrl.OnAction = caption
            ctrl.Visible = True
            Err.Clear
            Exit Sub
        End If
    Next ctrl

    Set btn = bar.Controls.Add(1)
    If btn Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    btn.Style = 2
    btn.Tag = "SendCheckScale"
    btn.Caption = caption
    btn.OnAction = caption
    btn.Visible = True
    Err.Clear

End Sub


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

    SendCheckEnsureButton

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
