Attribute VB_Name = "modAttachmentCheck"
Option Explicit

'========================================================
' 添付ファイルの中身を調べる
'
' 拡張子は見ない。先頭と末尾の ZIP 署名で判定し、
' セントラルディレクトリに書いてあるファイル名から拡張子を集める。
' .zi_ や ._IP のように拡張子が壊されていても、中身が ZIP なら読める。
'
' パスワードで中身が暗号化されていても、ファイル名自体が
' 暗号化されていなければ拡張子は取れる。
' xlsx / docx / pptx も実体は ZIP なので、渡すと中の拡張子が出る。
' 中まで読めた入れ子 ZIP は器なので、extensions には入れない。
'
' 戻り値の例:
'   file: C:\temp\見積._IP
'   isZip: True
'   outerExtension: ._ip
'   entryCount: 2
'   extensions: .pdf, .xlsm
'   entries:
'   見積.pdf    .pdf
'   inner.zip/名簿.xlsm    .xlsm
'
' 32MB を超えるファイル、分割 ZIP、ZIP64 は未対応。
' 入れ子は 4 段まで。展開が必要な入れ子は 1 ファイル 8MB まで。
'========================================================

Private Const MAX_FILE_BYTES As Long = 33554432
Private Const MAX_NESTED_BYTES As Long = 8388608
Private Const MAX_ENTRIES As Long = 2000
Private Const MAX_DEPTH As Long = 4
Private Const MAX_INFLATE As Long = 30

Private mFso As Object
Private mTempSeq As Long
Private mInflateCount As Long


'========================================================
' 公開
'========================================================

Public Function InspectAttachmentFile(ByVal filePath As String) As String

    Dim data() As Byte
    Dim lines As Collection
    Dim errors As Collection
    Dim exts As Object
    Dim isZip As Boolean

    On Error GoTo Fail

    mInflateCount = 0

    If Fso().FolderExists(filePath) Then
        InspectAttachmentFile = ReportError(filePath, "フォルダは調べられません")
        Exit Function
    End If

    If Not Fso().FileExists(filePath) Then
        InspectAttachmentFile = ReportError(filePath, "ファイルがありません")
        Exit Function
    End If

    If CDbl(Fso().GetFile(filePath).Size) = 0 Then
        InspectAttachmentFile = ReportPlainFile(filePath)
        Exit Function
    End If

    If CDbl(Fso().GetFile(filePath).Size) > CDbl(MAX_FILE_BYTES) Then
        InspectAttachmentFile = ReportError(filePath, "32MBを超えるため未走査です")
        Exit Function
    End If

    data = ReadFileBytes(filePath)

    Set lines = New Collection
    Set errors = New Collection
    Set exts = CreateObject("Scripting.Dictionary")

    isZip = WalkZip(data, "", 0, lines, exts, errors)

    If Not isZip Then
        InspectAttachmentFile = ReportPlainFile(filePath)
    Else
        InspectAttachmentFile = ReportZip(filePath, lines, exts, errors)
    End If

    Exit Function

Fail:
    InspectAttachmentFile = ReportError(filePath, Err.Description)

End Function


' extensions 行だけを返す。ZIP でなければ、そのファイル自身の拡張子。
Public Function AttachmentExtensions(ByVal filePath As String) As String

    AttachmentExtensions = LineValue(InspectAttachmentFile(filePath), "extensions: ")

End Function


'========================================================
' ZIP
'========================================================

Private Function WalkZip( _
    ByRef data() As Byte, _
    ByVal prefix As String, _
    ByVal depth As Long, _
    ByVal lines As Collection, _
    ByVal exts As Object, _
    ByVal errors As Collection) As Boolean

    Dim eocd As Long
    Dim diskNo As Long
    Dim cdDisk As Long
    Dim total As Long
    Dim cdSize As Long
    Dim cdOffset As Long
    Dim pos As Long
    Dim i As Long
    Dim flags As Long
    Dim method As Long
    Dim compSize As Long
    Dim uncompSize As Long
    Dim nameLen As Long
    Dim extraLen As Long
    Dim commentLen As Long
    Dim localOffset As Long
    Dim name As String
    Dim fullName As String
    Dim encrypted As Boolean
    Dim isDir As Boolean
    Dim openedAsZip As Boolean
    Dim errorsBefore As Long
    Dim sizesOk As Boolean
    Dim payload() As Byte

    If ByteCount(data) < 22 Then Exit Function
    If depth > MAX_DEPTH Then Exit Function

    eocd = FindEocd(data)
    If eocd < 0 Then Exit Function

    diskNo = ReadU16(data, eocd + 4)
    cdDisk = ReadU16(data, eocd + 6)
    total = ReadU16(data, eocd + 10)

    If diskNo <> 0 Or cdDisk <> 0 Then
        errors.Add "分割されたZIPには対応していません"
        WalkZip = True
        Exit Function
    End If

    If total = &HFFFF Or IsAllFF(data, eocd + 12) Or IsAllFF(data, eocd + 16) Then
        errors.Add "ZIP64には対応していません"
        WalkZip = True
        Exit Function
    End If

    If Not TryReadU32(data, eocd + 12, cdSize) Then Exit Function
    If Not TryReadU32(data, eocd + 16, cdOffset) Then Exit Function

    If total = 0 And cdSize = 0 Then
        WalkZip = True
        Exit Function
    End If

    If Not IsSig(data, cdOffset, &H50, &H4B, &H1, &H2) Then Exit Function

    WalkZip = True
    pos = cdOffset

    For i = 1 To total

        If lines.Count >= MAX_ENTRIES Then
            errors.Add "エントリ数が上限を超えたため打ち切りました"
            Exit For
        End If

        If Not InRange(data, pos, 46) Then
            errors.Add "セントラルディレクトリが不正です"
            Exit For
        End If

        If Not IsSig(data, pos, &H50, &H4B, &H1, &H2) Then
            errors.Add "セントラルディレクトリが不正です"
            Exit For
        End If

        flags = ReadU16(data, pos + 8)
        method = ReadU16(data, pos + 10)
        nameLen = ReadU16(data, pos + 28)
        extraLen = ReadU16(data, pos + 30)
        commentLen = ReadU16(data, pos + 32)

        If Not InRange(data, pos + 46, nameLen + extraLen + commentLen) Then
            errors.Add "セントラルディレクトリが不正です"
            Exit For
        End If

        name = DecodeName(data, pos + 46, nameLen, (flags And &H800) <> 0)

        If (flags And &H800) = 0 And extraLen > 0 Then
            name = PreferUnicodePath(data, pos + 46 + nameLen, extraLen, name)
        End If

        name = NormalizeName(name)

        sizesOk = False
        compSize = 0
        uncompSize = 0
        localOffset = 0

        If Not IsAllFF(data, pos + 20) And Not IsAllFF(data, pos + 24) And Not IsAllFF(data, pos + 42) Then
            If TryReadU32(data, pos + 20, compSize) Then
                If TryReadU32(data, pos + 24, uncompSize) Then
                    If TryReadU32(data, pos + 42, localOffset) Then
                        sizesOk = True
                    End If
                End If
            End If
        End If

        isDir = (Len(name) > 0 And Right$(name, 1) = "/")
        encrypted = ((flags And &H1) <> 0)

        If Len(name) > 0 And Not isDir Then

            If Len(prefix) = 0 Then
                fullName = name
            Else
                fullName = prefix & name
            End If

            openedAsZip = False

            If Not encrypted And sizesOk And compSize > 0 And compSize <= MAX_NESTED_BYTES Then
                If depth < MAX_DEPTH And (method = 0 Or (method = 8 And LooksLikeZipName(name))) Then
                    If TryGetPayload(data, localOffset, compSize, method, payload) Then
                        If HasZipMagic(payload) And ByteCount(payload) <= MAX_NESTED_BYTES Then
                            errorsBefore = errors.Count
                            If WalkZip(payload, fullName & "/", depth + 1, lines, exts, errors) Then
                                If errors.Count = errorsBefore Then openedAsZip = True
                            Else
                                errors.Add "中のZIPを読めません: " & fullName
                            End If
                        End If
                    ElseIf method = 8 And LooksLikeZipName(name) Then
                        If mInflateCount > MAX_INFLATE Then
                            errors.Add "展開の上限を超えたため未走査です: " & fullName
                        Else
                            errors.Add "中のZIPを展開できません: " & fullName
                        End If
                    End If
                ElseIf depth >= MAX_DEPTH And LooksLikeZipName(name) Then
                    errors.Add "入れ子が深すぎるため未走査です: " & fullName
                End If
            End If

            AddFileEntry lines, exts, fullName, encrypted, openedAsZip

        End If

        pos = pos + 46 + nameLen + extraLen + commentLen

    Next i

End Function


Private Function TryGetPayload( _
    ByRef data() As Byte, _
    ByVal localOffset As Long, _
    ByVal compSize As Long, _
    ByVal method As Long, _
    ByRef payload() As Byte) As Boolean

    Dim nameLen As Long
    Dim extraLen As Long
    Dim dataStart As Long
    Dim startAt As Double
    Dim compressed() As Byte

    If method <> 0 And method <> 8 Then Exit Function
    If localOffset < LBound(data) Then Exit Function
    If Not IsSig(data, localOffset, &H50, &H4B, &H3, &H4) Then Exit Function
    If Not InRange(data, localOffset, 30) Then Exit Function

    nameLen = ReadU16(data, localOffset + 26)
    extraLen = ReadU16(data, localOffset + 28)
    startAt = CDbl(localOffset) + 30# + CDbl(nameLen) + CDbl(extraLen)

    If startAt + CDbl(compSize) - 1# > CDbl(UBound(data)) Then Exit Function
    If startAt > 2147483647# Then Exit Function

    dataStart = CLng(startAt)
    compressed = SliceBytes(data, dataStart, compSize)

    If method = 0 Then
        payload = compressed
        TryGetPayload = True
    Else
        TryGetPayload = InflateDeflate(compressed, payload)
    End If

End Function


Private Function InflateDeflate( _
    ByRef compressed() As Byte, _
    ByRef payload() As Byte) As Boolean

    Dim inPath As String
    Dim outPath As String
    Dim ps1 As String
    Dim sh As Object
    Dim code As Long

    On Error GoTo Fail

    mInflateCount = mInflateCount + 1
    If mInflateCount > MAX_INFLATE Then GoTo Fail

    inPath = TempFilePath(".deflate")
    outPath = TempFilePath(".bin")
    ps1 = TempFilePath(".ps1")

    WriteBytes inPath, compressed
    WriteAsciiFile ps1, InflateScript()

    Set sh = CreateObject("WScript.Shell")
    code = sh.Run( _
        "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & _
        ps1 & """ -inPath """ & inPath & """ -outPath """ & outPath & """", _
        0, True)

    If code <> 0 Then GoTo Fail
    If Not Fso().FileExists(outPath) Then GoTo Fail
    If CDbl(Fso().GetFile(outPath).Size) <= 0 Then GoTo Fail
    If CDbl(Fso().GetFile(outPath).Size) > CDbl(MAX_NESTED_BYTES) Then GoTo Fail

    payload = ReadFileBytes(outPath)
    InflateDeflate = True
    GoTo Cleanup

Fail:
    InflateDeflate = False

Cleanup:
    On Error Resume Next
    DeleteFileIfExists inPath
    DeleteFileIfExists outPath
    DeleteFileIfExists ps1

End Function


Private Function InflateScript() As String

    InflateScript = _
        "param([string]$inPath, [string]$outPath)" & vbCrLf & _
        "$ErrorActionPreference = 'Stop'" & vbCrLf & _
        "Add-Type -AssemblyName System.IO.Compression" & vbCrLf & _
        "$inputStream = [System.IO.File]::OpenRead($inPath)" & vbCrLf & _
        "try {" & vbCrLf & _
        "  $deflate = New-Object System.IO.Compression.DeflateStream($inputStream, [System.IO.Compression.CompressionMode]::Decompress)" & vbCrLf & _
        "  try {" & vbCrLf & _
        "    $outputStream = [System.IO.File]::Create($outPath)" & vbCrLf & _
        "    try { $deflate.CopyTo($outputStream) } finally { $outputStream.Dispose() }" & vbCrLf & _
        "  } finally { $deflate.Dispose() }" & vbCrLf & _
        "} finally { $inputStream.Dispose() }" & vbCrLf

End Function


'========================================================
' レポート
'========================================================

Private Function ReportZip( _
    ByVal filePath As String, _
    ByVal lines As Collection, _
    ByVal exts As Object, _
    ByVal errors As Collection) As String

    Dim s As String

    s = HeaderText(filePath, True, lines.Count, ExtensionsText(exts))
    s = s & "entries:" & vbCrLf
    s = s & CollectionText(lines)
    s = s & ErrorsText(errors)
    ReportZip = s

End Function


Private Function ReportPlainFile(ByVal filePath As String) As String

    Dim lines As Collection
    Dim exts As Object
    Dim name As String

    Set lines = New Collection
    Set exts = CreateObject("Scripting.Dictionary")
    name = Fso().GetFileName(filePath)
    AddFileEntry lines, exts, name, False

    ReportPlainFile = HeaderText(filePath, False, 1, ExtensionsText(exts)) & _
        "entries:" & vbCrLf & _
        CollectionText(lines)

End Function


Private Function ReportError(ByVal filePath As String, ByVal message As String) As String

    ReportError = HeaderText(filePath, False, 0, "") & _
        "entries:" & vbCrLf & _
        "errors:" & vbCrLf & _
        message & vbCrLf

End Function


Private Function HeaderText( _
    ByVal filePath As String, _
    ByVal isZip As Boolean, _
    ByVal entryCount As Long, _
    ByVal extensions As String) As String

    Dim zipText As String

    If isZip Then
        zipText = "True"
    Else
        zipText = "False"
    End If

    HeaderText = _
        "file: " & filePath & vbCrLf & _
        "isZip: " & zipText & vbCrLf & _
        "outerExtension: " & DisplayExt(Fso().GetFileName(filePath)) & vbCrLf & _
        "entryCount: " & CStr(entryCount) & vbCrLf & _
        "extensions: " & extensions & vbCrLf

End Function


Private Function ErrorsText(ByVal errors As Collection) As String

    Dim i As Long

    If errors.Count = 0 Then Exit Function

    ErrorsText = "errors:" & vbCrLf
    For i = 1 To errors.Count
        ErrorsText = ErrorsText & CStr(errors(i)) & vbCrLf
    Next i

End Function


Private Sub AddFileEntry( _
    ByVal lines As Collection, _
    ByVal exts As Object, _
    ByVal fullName As String, _
    ByVal encrypted As Boolean, _
    Optional ByVal isContainer As Boolean = False)

    Dim ext As String
    Dim line As String

    ext = DisplayExt(fullName)
    line = fullName & vbTab & ext
    If isContainer Then line = line & vbTab & "container"
    If encrypted Then line = line & vbTab & "encrypted"
    lines.Add line

    If Not isContainer Then
        If Not exts.Exists(ext) Then exts.Add ext, True
    End If

End Sub


Private Function ExtensionsText(ByVal exts As Object) As String

    Dim arr() As String
    Dim keys As Variant
    Dim i As Long

    If exts.Count = 0 Then Exit Function

    ReDim arr(0 To exts.Count - 1)
    keys = exts.Keys
    For i = 0 To exts.Count - 1
        arr(i) = CStr(keys(i))
    Next i

    SortStrings arr
    ExtensionsText = Join(arr, ", ")

End Function


Private Function CollectionText(ByVal lines As Collection) As String

    Dim arr() As String
    Dim i As Long

    If lines.Count = 0 Then Exit Function

    ReDim arr(0 To lines.Count - 1)
    For i = 1 To lines.Count
        arr(i - 1) = CStr(lines(i))
    Next i

    SortStrings arr

    For i = 0 To UBound(arr)
        CollectionText = CollectionText & arr(i) & vbCrLf
    Next i

End Function


'========================================================
' 名前と拡張子
'========================================================

Private Function ExtensionOf(ByVal relativePath As String) As String

    Dim base As String
    Dim slash As Long
    Dim dot As Long

    base = Replace$(relativePath, "\", "/")
    slash = InStrRev(base, "/")
    If slash > 0 Then base = Mid$(base, slash + 1)

    dot = InStrRev(base, ".")
    If dot <= 1 Then Exit Function

    ExtensionOf = LCase$(Mid$(base, dot))

End Function


Private Function DisplayExt(ByVal relativePath As String) As String

    DisplayExt = ExtensionOf(relativePath)
    If Len(DisplayExt) = 0 Then DisplayExt = "(none)"

End Function


Private Function LooksLikeZipName(ByVal relativePath As String) As Boolean

    Dim ext As String

    ext = ExtensionOf(relativePath)
    If Len(ext) = 0 Then Exit Function
    ext = Mid$(ext, 2)

    Select Case ext
        Case "zip", "zipx", "zip_", "zi_", "z_p", "_ip"
            LooksLikeZipName = True
    End Select

End Function


Private Function NormalizeName(ByVal name As String) As String

    name = Replace$(name, "\", "/")
    Do While Left$(name, 1) = "/"
        name = Mid$(name, 2)
    Loop
    NormalizeName = name

End Function


Private Function DecodeName( _
    ByRef data() As Byte, _
    ByVal pos As Long, _
    ByVal length As Long, _
    ByVal utf8 As Boolean) As String

    If length <= 0 Then Exit Function

    If utf8 Then
        DecodeName = DecodeUtf8(data, pos, length)
    Else
        DecodeName = DecodeAnsi(data, pos, length)
    End If

End Function


' Info-ZIP の Unicode Path (0x7075) があれば、そちらをファイル名にする。
Private Function PreferUnicodePath( _
    ByRef data() As Byte, _
    ByVal extraPos As Long, _
    ByVal extraLen As Long, _
    ByVal currentName As String) As String

    Dim p As Long
    Dim endP As Long
    Dim id As Long
    Dim sz As Long

    PreferUnicodePath = currentName
    If extraLen < 4 Then Exit Function

    p = extraPos
    endP = extraPos + extraLen

    Do While p + 4 <= endP
        id = ReadU16(data, p)
        sz = ReadU16(data, p + 2)
        If p + 4 + sz > endP Then Exit Function

        If id = &H7075 And sz >= 6 Then
            If data(p + 4) = 1 Then
                PreferUnicodePath = DecodeUtf8(data, p + 9, sz - 5)
                Exit Function
            End If
        End If

        p = p + 4 + sz
    Loop

End Function


Private Function DecodeUtf8( _
    ByRef data() As Byte, _
    ByVal pos As Long, _
    ByVal length As Long) As String

    Dim slice() As Byte
    Dim stm As Object

    On Error GoTo Fallback

    slice = SliceBytes(data, pos, length)
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write slice
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "utf-8"
    DecodeUtf8 = stm.ReadText
    stm.Close
    Exit Function

Fallback:
    DecodeUtf8 = DecodeAnsi(data, pos, length)

End Function


Private Function DecodeAnsi( _
    ByRef data() As Byte, _
    ByVal pos As Long, _
    ByVal length As Long) As String

    Dim slice() As Byte

    slice = SliceBytes(data, pos, length)
    DecodeAnsi = StrConv(slice, vbUnicode)

End Function


'========================================================
' バイナリ
'========================================================

Private Function FindEocd(ByRef data() As Byte) As Long

    Dim i As Long
    Dim minPos As Long
    Dim commentLen As Long

    FindEocd = -1
    If ByteCount(data) < 22 Then Exit Function

    minPos = LBound(data)
    If ByteCount(data) > 65557 Then minPos = UBound(data) - 65556

    For i = UBound(data) - 21 To minPos Step -1
        If IsSig(data, i, &H50, &H4B, &H5, &H6) Then
            commentLen = ReadU16(data, i + 20)
            If i + 22 + commentLen - 1 = UBound(data) Then
                FindEocd = i
                Exit Function
            End If
        End If
    Next i

End Function


Private Function HasZipMagic(ByRef data() As Byte) As Boolean

    Dim i As Long

    If ByteCount(data) < 4 Then Exit Function
    i = LBound(data)
    If data(i) <> &H50 Or data(i + 1) <> &H4B Then Exit Function

    If data(i + 2) = &H3 And data(i + 3) = &H4 Then HasZipMagic = True
    If data(i + 2) = &H5 And data(i + 3) = &H6 Then HasZipMagic = True
    If data(i + 2) = &H7 And data(i + 3) = &H8 Then HasZipMagic = True

End Function


Private Function IsSig( _
    ByRef data() As Byte, _
    ByVal pos As Long, _
    ByVal b0 As Byte, _
    ByVal b1 As Byte, _
    ByVal b2 As Byte, _
    ByVal b3 As Byte) As Boolean

    If pos < LBound(data) Then Exit Function
    If pos + 3 > UBound(data) Then Exit Function

    IsSig = _
        data(pos) = b0 And _
        data(pos + 1) = b1 And _
        data(pos + 2) = b2 And _
        data(pos + 3) = b3

End Function


Private Function IsAllFF(ByRef data() As Byte, ByVal pos As Long) As Boolean

    If Not InRange(data, pos, 4) Then Exit Function

    IsAllFF = _
        data(pos) = 255 And _
        data(pos + 1) = 255 And _
        data(pos + 2) = 255 And _
        data(pos + 3) = 255

End Function


Private Function ReadU16(ByRef data() As Byte, ByVal pos As Long) As Long

    ReadU16 = CLng(data(pos)) + CLng(data(pos + 1)) * 256&

End Function


Private Function TryReadU32( _
    ByRef data() As Byte, _
    ByVal pos As Long, _
    ByRef value As Long) As Boolean

    Dim v As Double

    If Not InRange(data, pos, 4) Then Exit Function

    v = CDbl(data(pos)) + _
        CDbl(data(pos + 1)) * 256# + _
        CDbl(data(pos + 2)) * 65536# + _
        CDbl(data(pos + 3)) * 16777216#

    If v > 2147483647# Then Exit Function

    value = CLng(v)
    TryReadU32 = True

End Function


Private Function InRange( _
    ByRef data() As Byte, _
    ByVal pos As Long, _
    ByVal size As Long) As Boolean

    If size < 0 Then Exit Function
    If pos < LBound(data) Then Exit Function
    If pos + size - 1 > UBound(data) Then Exit Function
    InRange = True

End Function


Private Function ByteCount(ByRef data() As Byte) As Long

    On Error GoTo EmptyArr
    ByteCount = UBound(data) - LBound(data) + 1
    Exit Function

EmptyArr:
    ByteCount = 0

End Function


Private Function SliceBytes( _
    ByRef data() As Byte, _
    ByVal pos As Long, _
    ByVal length As Long) As Byte()

    Dim out() As Byte
    Dim i As Long

    ReDim out(0 To length - 1)
    For i = 0 To length - 1
        out(i) = data(pos + i)
    Next i
    SliceBytes = out

End Function


Private Function ReadFileBytes(ByVal filePath As String) As Byte()

    Dim fn As Integer
    Dim n As Long
    Dim data() As Byte

    fn = FreeFile
    Open filePath For Binary Access Read As #fn
    n = LOF(fn)
    If n > 0 Then
        ReDim data(0 To n - 1)
        Get #fn, 1, data
    End If
    Close #fn
    ReadFileBytes = data

End Function


Private Sub WriteBytes(ByVal filePath As String, ByRef data() As Byte)

    Dim fn As Integer

    DeleteFileIfExists filePath
    fn = FreeFile
    Open filePath For Binary Access Write As #fn
    Put #fn, 1, data
    Close #fn

End Sub


'========================================================
' ファイル
'========================================================

Private Function Fso() As Object

    If mFso Is Nothing Then
        Set mFso = CreateObject("Scripting.FileSystemObject")
    End If
    Set Fso = mFso

End Function


Private Function LineValue(ByVal report As String, ByVal prefix As String) As String

    Dim rows() As String
    Dim i As Long

    rows = Split(report, vbCrLf)
    For i = LBound(rows) To UBound(rows)
        If Left$(rows(i), Len(prefix)) = prefix Then
            LineValue = Mid$(rows(i), Len(prefix) + 1)
            Exit Function
        End If
    Next i

End Function


Private Sub SortStrings(ByRef items() As String)

    Dim i As Long
    Dim j As Long
    Dim tmp As String

    If UBound(items) <= LBound(items) Then Exit Sub

    For i = LBound(items) To UBound(items) - 1
        For j = i + 1 To UBound(items)
            If StrComp(items(j), items(i), vbTextCompare) < 0 Then
                tmp = items(i)
                items(i) = items(j)
                items(j) = tmp
            End If
        Next j
    Next i

End Sub


Private Function TempFilePath(ByVal ext As String) As String

    Dim folder As String

    mTempSeq = mTempSeq + 1
    folder = Fso().GetSpecialFolder(2).Path & "\SendCheckAttach"
    If Not Fso().FolderExists(folder) Then Fso().CreateFolder folder
    TempFilePath = folder & "\p" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(mTempSeq) & ext

End Function


Private Sub WriteAsciiFile(ByVal filePath As String, ByVal text As String)

    Dim ts As Object

    Set ts = Fso().CreateTextFile(filePath, True, False)
    ts.Write text
    ts.Close

End Sub


Private Sub DeleteFileIfExists(ByVal filePath As String)

    On Error Resume Next
    If Len(filePath) = 0 Then Exit Sub
    If Fso().FileExists(filePath) Then Fso().DeleteFile filePath, True

End Sub
