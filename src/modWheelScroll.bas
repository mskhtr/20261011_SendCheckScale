Attribute VB_Name = "modWheelScroll"
Option Explicit

'========================================================
' マウスホイールで Frame をスクロールする
'
' UserForm の Frame は、標準ではホイールに反応しない。
' フォーム表示中だけフックを掛け、カーソル下の縦スクロール
' 付き Frame の ScrollTop を動かす。
'========================================================

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0

' 宛先カード 1 枚ぶん
Private Const WHEEL_STEP As Single = 100

Private mHook As LongPtr
Private mForm As Object
Private mInHook As Boolean

Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" ( _
    ByVal idHook As Long, _
    ByVal lpfn As LongPtr, _
    ByVal hMod As LongPtr, _
    ByVal dwThreadId As Long) As LongPtr

Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hhk As LongPtr) As Long

Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
    ByVal hhk As LongPtr, _
    ByVal nCode As Long, _
    ByVal wParam As LongPtr, _
    ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function FindWindowW Lib "user32" ( _
    ByVal lpClassName As LongPtr, _
    ByVal lpWindowName As LongPtr) As LongPtr

Private Declare PtrSafe Function ClientToScreen Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    lpPoint As POINTAPI) As Long

Private Declare PtrSafe Function GetClientRect Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    lpRect As RECT) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByVal Source As LongPtr, _
    ByVal Length As LongPtr)

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Sub WheelScroll_Attach(ByVal frm As Object)

    Set mForm = frm

    If mHook <> 0 Then Exit Sub

    mHook = SetWindowsHookEx( _
        WH_MOUSE_LL, _
        AddressOf WheelScroll_HookProc, _
        0, _
        0)

End Sub


Public Sub WheelScroll_Detach()

    If mHook <> 0 Then
        UnhookWindowsHookEx mHook
        mHook = 0
    End If

    Set mForm = Nothing

End Sub


Private Function WheelScroll_HookProc( _
    ByVal nCode As Long, _
    ByVal wParam As LongPtr, _
    ByVal lParam As LongPtr) As LongPtr

    On Error GoTo PassThrough

    If mInHook Then GoTo PassThrough
    If nCode <> HC_ACTION Then GoTo PassThrough
    If wParam <> WM_MOUSEWHEEL Then GoTo PassThrough
    If mForm Is Nothing Then GoTo PassThrough

    mInHook = True

    If ScrollFrameUnderCursor(lParam) Then
        mInHook = False
        WheelScroll_HookProc = 1
        Exit Function
    End If

    mInHook = False

PassThrough:
    mInHook = False
    WheelScroll_HookProc = CallNextHookEx(mHook, nCode, wParam, lParam)

End Function


Private Function ScrollFrameUnderCursor(ByVal lParam As LongPtr) As Boolean

    Dim hwnd As LongPtr
    Dim origin As POINTAPI
    Dim rc As RECT
    Dim scaleX As Single
    Dim scaleY As Single
    Dim cursorX As Long
    Dim cursorY As Long
    Dim delta As Integer
    Dim ptX As Single
    Dim ptY As Single
    Dim clsName As String
    Dim cap As String
    Dim ctl As Object
    Dim target As Object
    Dim newTop As Single
    Dim maxTop As Single
    Dim notches As Single
    Dim stepSize As Single

    On Error GoTo Fail

    clsName = "ThunderDFrame"
    cap = mForm.Caption
    hwnd = FindWindowW(StrPtr(clsName), StrPtr(cap))
    If hwnd = 0 Then Exit Function

    If GetClientRect(hwnd, rc) = 0 Then Exit Function
    If mForm.InsideWidth <= 0 Then Exit Function
    If mForm.InsideHeight <= 0 Then Exit Function

    scaleX = rc.Right / mForm.InsideWidth
    scaleY = rc.Bottom / mForm.InsideHeight
    If scaleX <= 0 Or scaleY <= 0 Then Exit Function

    origin.X = 0
    origin.Y = 0
    ClientToScreen hwnd, origin

    CopyMemory cursorX, lParam, 4
    CopyMemory cursorY, lParam + 4, 4
    CopyMemory delta, lParam + 10, 2

    If cursorX < origin.X Or cursorX >= origin.X + rc.Right Then Exit Function
    If cursorY < origin.Y Or cursorY >= origin.Y + rc.Bottom Then Exit Function
    If delta = 0 Then Exit Function

    ptX = (cursorX - origin.X) / scaleX
    ptY = (cursorY - origin.Y) / scaleY
    notches = CSng(delta) / 120

    For Each ctl In mForm.Controls

        If TypeName(ctl) = "Frame" Then
            If ctl.Visible Then
            If PointInControl(ctl, ptX, ptY) Then

                Set target = FindVerticalScrollFrame(ctl)
                If Not target Is Nothing Then

                    If target.Name = "fraAttachments_list" Then
                        stepSize = 22
                    Else
                        stepSize = WHEEL_STEP
                    End If

                    maxTop = target.ScrollHeight - target.InsideHeight
                    If maxTop < 0 Then maxTop = 0

                    newTop = target.ScrollTop - notches * stepSize
                    If newTop < 0 Then newTop = 0
                    If newTop > maxTop Then newTop = maxTop

                    target.ScrollTop = newTop
                    On Error Resume Next
                    mForm.SyncScrollThumb target
                    On Error GoTo Fail
                    ScrollFrameUnderCursor = True
                    Exit Function

                End If
            End If
            End If
        End If

    Next ctl

    Exit Function

Fail:
    ScrollFrameUnderCursor = False

End Function


Private Function FindVerticalScrollFrame(ByVal container As Object) As Object

    Dim ctl As Object

    On Error GoTo Fail

    If TypeName(container) = "Frame" Then
        If Right$(container.Name, 5) = "_list" Then
            If container.ScrollHeight > container.InsideHeight + 1 Then
                Set FindVerticalScrollFrame = container
                Exit Function
            End If
        End If
    End If

    For Each ctl In container.Controls
        If TypeName(ctl) = "Frame" Then
            Set FindVerticalScrollFrame = FindVerticalScrollFrame(ctl)
            If Not FindVerticalScrollFrame Is Nothing Then Exit Function
        End If
    Next ctl

    Exit Function

Fail:
    Set FindVerticalScrollFrame = Nothing

End Function


Private Function PointInControl( _
    ByVal ctl As Object, _
    ByVal ptX As Single, _
    ByVal ptY As Single) As Boolean

    PointInControl = _
        ptX >= ctl.Left And ptX < ctl.Left + ctl.Width And _
        ptY >= ctl.Top And ptY < ctl.Top + ctl.Height

End Function
