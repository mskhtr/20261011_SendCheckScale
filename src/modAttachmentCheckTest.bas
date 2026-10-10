Attribute VB_Name = "modAttachmentCheckTest"
Option Explicit

'========================================================
' 添付チェックの動作確認
'
' Excel で modAttachmentCheck とこのモジュールを読み込み、
' 次を実行する。
'
'   RunAttachmentCheckTests "C:\path\to\20261008_SendCheck\test"
'
' 結果は test\result.txt に出る。
' テスト用 ZIP は test\fixtures に残す。
'========================================================

Private Type ZipMeta
    Name As String
    ForceUtf8 As Boolean
    Encrypted As Boolean
    Crc As Long
    Length As Long
    Offset As Long
End Type

Private Type Buf
    Data() As Byte
    Length As Long
End Type

Private mFso As Object
Private mPass As Long
Private mFail As Long
Private mLog As String
Private mSample As String
Private mMeta() As ZipMeta
Private mPartCount As Long
Private mLocal As Buf


Public Function RunAttachmentCheckTests(ByVal testRoot As String) As String

    On Error GoTo Fail

    mPass = 0
    mFail = 0
    mLog = ""

    If Not Fso().FolderExists(testRoot) Then Fso().CreateFolder testRoot

    CheckCrc
    CheckTextRoundtrip
    CheckPlainFiles testRoot
    CheckStoredZip testRoot
    CheckContainerExcluded testRoot
    CheckDeflateZip testRoot

    mLog = mLog & vbCrLf & CStr(mPass) & " passed, " & CStr(mFail) & " failed" & vbCrLf
    If Len(mSample) > 0 Then
        mLog = mLog & vbCrLf & "---- sample: stored._IP ----" & vbCrLf & mSample
    End If
    WriteUtf8File testRoot & "\result.txt", mLog
    RunAttachmentCheckTests = mLog
    Exit Function

Fail:
    mLog = mLog & vbCrLf & "FAIL runner" & vbCrLf & _
        CStr(Err.Number) & " " & Err.Description & vbCrLf
    On Error Resume Next
    WriteUtf8File testRoot & "\result.txt", mLog
    RunAttachmentCheckTests = mLog

End Function


'========================================================
' ケース
'========================================================

Private Sub CheckCrc()

    Dim sample() As Byte
    Dim text As String
    Dim i As Long

    text = "123456789"
    ReDim sample(0 To Len(text) - 1)
    For i = 1 To Len(text)
        sample(i - 1) = Asc(Mid$(text, i, 1))
    Next i

    If U32Hex(Crc32(sample, ByteCount(sample))) = "CBF43926" Then
        Pass "CRC32"
    Else
        Fail "CRC32", U32Hex(Crc32(sample, ByteCount(sample)))
    End If

End Sub


Private Sub CheckTextRoundtrip()

    Dim encoded() As Byte
    Dim back As String

    encoded = Utf8Bytes(JpMemoName())
    back = Utf8Text(encoded)
    If back = JpMemoName() Then
        Pass "UTF-8の往復"
    Else
        Fail "UTF-8の往復", back
    End If

    encoded = AnsiBytes(JpEstimateName())
    back = StrConv(encoded, vbUnicode)
    If back = JpEstimateName() Then
        Pass "Shift_JISの往復"
    Else
        Fail "Shift_JISの往復", back
    End If

End Sub


Private Sub CheckPlainFiles(ByVal testRoot As String)

    Dim dirPath As String
    Dim plainPath As String
    Dim fakeZip As String
    Dim pkPath As String
    Dim report As String

    dirPath = testRoot & "\fixtures"
    If Fso().FolderExists(dirPath) Then Fso().DeleteFolder dirPath, True
    Fso().CreateFolder dirPath

    plainPath = dirPath & "\note.txt"
    fakeZip = dirPath & "\fake.zip"
    pkPath = dirPath & "\pk.txt"

    WriteAsciiFile plainPath, "hello"
    WriteAsciiFile fakeZip, "this is not a zip"
    WriteAsciiFile pkPath, "PK this is not a zip"

    report = InspectAttachmentFile(plainPath)
    If LineOf(report, "isZip: ") = "False" And LineOf(report, "extensions: ") = ".txt" And Not HasErrors(report) Then
        Pass "通常のテキスト"
    Else
        Fail "通常のテキスト", report
    End If

    report = InspectAttachmentFile(fakeZip)
    If LineOf(report, "isZip: ") = "False" And LineOf(report, "extensions: ") = ".zip" And Not HasErrors(report) Then
        Pass "拡張子だけzip"
    Else
        Fail "拡張子だけzip", report
    End If

    report = InspectAttachmentFile(pkPath)
    If LineOf(report, "isZip: ") = "False" And Not HasErrors(report) Then
        Pass "PKで始まるテキスト"
    Else
        Fail "PKで始まるテキスト", report
    End If

    report = InspectAttachmentFile(dirPath & "\missing.bin")
    If InStr(report, "ファイルがありません") > 0 Then
        Pass "存在しないファイル"
    Else
        Fail "存在しないファイル", report
    End If

End Sub


Private Sub CheckStoredZip(ByVal testRoot As String)

    Dim dirPath As String
    Dim innerPath As String
    Dim storedPath As String
    Dim ziPath As String
    Dim ipPath As String
    Dim ansiPath As String
    Dim encryptedPath As String
    Dim emptyPath As String
    Dim innerBytes() As Byte
    Dim emptyBytes(0 To 21) As Byte
    Dim report As String
    Dim renamedZi As String
    Dim renamedIp As String
    Dim dotNetCode As Long

    dirPath = testRoot & "\fixtures"
    innerPath = dirPath & "\inner-stored.zip"
    storedPath = dirPath & "\stored.zip"
    ziPath = dirPath & "\stored.zi_"
    ipPath = dirPath & "\stored._IP"
    ansiPath = dirPath & "\ansi.zip"
    encryptedPath = dirPath & "\encrypted.zip"
    emptyPath = dirPath & "\empty.zip"

    ClearParts
    AddTextPart "secret.xlsm", "secret", False, False
    SaveParts innerPath
    innerBytes = ReadFileBytes(innerPath)

    ClearParts
    AddTextPart "a.txt", "alpha", False, False
    AddTextPart "b.PDF", "pdf", False, False
    AddTextPart "folder/sub.txt", "sub", False, False
    AddTextPart "run.exe", "dummy", False, False
    AddTextPart "readme", "hello", False, False
    AddTextPart JpMemoName(), "utf8", True, False
    AddRawPart "inner.zip", innerBytes, False, False
    AddRawPart "hidden.bin", innerBytes, False, False
    AddRawPart "mask.zi_", innerBytes, False, False
    SaveParts storedPath

    Fso().CopyFile storedPath, ziPath, True
    Fso().CopyFile storedPath, ipPath, True

    ClearParts
    AddTextPart JpEstimateName(), "ansi", False, False
    SaveParts ansiPath

    ClearParts
    AddTextPart "setup.exe", "dummy", False, True
    SaveParts encryptedPath

    emptyBytes(0) = &H50
    emptyBytes(1) = &H4B
    emptyBytes(2) = &H5
    emptyBytes(3) = &H6
    WriteBytes emptyPath, emptyBytes

    dotNetCode = RunPowerShellFile(dirPath & "\verify-stored.ps1", VerifyScript(), storedPath)
    If dotNetCode = 0 Then
        Pass "格納ZIPを.NETが読める"
    Else
        Fail "格納ZIPを.NETが読める", "exit " & CStr(dotNetCode)
    End If

    report = InspectAttachmentFile(emptyPath)
    If LineOf(report, "isZip: ") = "True" And LineOf(report, "entryCount: ") = "0" And Not HasErrors(report) Then
        Pass "空のZIP"
    Else
        Fail "空のZIP", report
    End If

    report = InspectAttachmentFile(storedPath)
    If Not HasErrors(report) And LineOf(report, "isZip: ") = "True" And LineOf(report, "entryCount: ") = "12" Then
        Pass "格納ZIPの件数"
    Else
        Fail "格納ZIPの件数", report
    End If

    ExpectExts report, "格納ZIPの拡張子", "(none)", ".exe", ".pdf", ".txt", ".xlsm"

    If CountText(report, "secret.xlsm") = 3 And InStr(report, JpMemoName()) > 0 And InStr(report, "hidden.bin/secret.xlsm") > 0 Then
        Pass "格納ZIPの入れ子と日本語名"
    Else
        Fail "格納ZIPの入れ子と日本語名", report
    End If

    If AttachmentExtensions(storedPath) = LineOf(report, "extensions: ") Then
        Pass "AttachmentExtensions"
    Else
        Fail "AttachmentExtensions", AttachmentExtensions(storedPath)
    End If

    renamedZi = InspectAttachmentFile(ziPath)
    renamedIp = InspectAttachmentFile(ipPath)

    If LineOf(renamedZi, "outerExtension: ") = ".zi_" And Not HasErrors(renamedZi) Then
        Pass "拡張子zi_"
    Else
        Fail "拡張子zi_", renamedZi
    End If

    If LineOf(renamedIp, "outerExtension: ") = "._ip" And Not HasErrors(renamedIp) Then
        Pass "拡張子_IP"
    Else
        Fail "拡張子_IP", renamedIp
    End If

    If EntriesOf(report) = EntriesOf(renamedZi) And EntriesOf(report) = EntriesOf(renamedIp) Then
        Pass "改名しても中身は同じ"
    Else
        Fail "改名しても中身は同じ", renamedZi & vbCrLf & renamedIp
    End If

    WriteUtf8File dirPath & "\report-renamed.txt", renamedIp
    mSample = renamedIp

    report = InspectAttachmentFile(ansiPath)
    If Not HasErrors(report) And InStr(report, JpEstimateName()) > 0 And LineOf(report, "extensions: ") = ".txt" Then
        Pass "日本語ファイル名Shift_JIS"
    Else
        Fail "日本語ファイル名Shift_JIS", report
    End If

    report = InspectAttachmentFile(encryptedPath)
    If Not HasErrors(report) And InStr(report, "setup.exe") > 0 And InStr(report, "encrypted") > 0 And LineOf(report, "extensions: ") = ".exe" Then
        Pass "暗号化フラグ"
    Else
        Fail "暗号化フラグ", report
    End If

End Sub


Private Sub CheckDeflateZip(ByVal testRoot As String)

    Dim dirPath As String
    Dim srcPath As String
    Dim packPath As String
    Dim zipPath As String
    Dim renamedPath As String
    Dim report As String
    Dim renamed As String
    Dim code As Long

    dirPath = testRoot & "\fixtures"
    srcPath = dirPath & "\deflate-src"
    packPath = srcPath & "\pack.zi_"
    zipPath = dirPath & "\real.zip"
    renamedPath = dirPath & "\real.ZI_"

    Fso().CreateFolder srcPath
    WriteAsciiFile srcPath & "\letter.docx", "doc"
    WriteAsciiFile srcPath & "\photo.png", "img"
    WriteAsciiFile srcPath & "\inside.xlsm", "xl"

    ClearParts
    AddTextPart "secret.xlsm", "secret", False, False
    SaveParts packPath

    code = RunPowerShellFile(dirPath & "\make-deflate.ps1", MakeDeflateScript(), srcPath & "|" & zipPath)
    If code <> 0 Then
        Fail "実圧縮ZIP", "zip作成に失敗 exit " & CStr(code)
        Exit Sub
    End If

    Fso().CopyFile zipPath, renamedPath, True

    report = InspectAttachmentFile(zipPath)
    If HasErrors(report) Or LineOf(report, "isZip: ") <> "True" Then
        Fail "実圧縮ZIP", report
        Exit Sub
    End If

    ExpectExts report, "実圧縮ZIPの拡張子", ".docx", ".png", ".txt", ".xlsm"

    If InStr(report, JpMemoName()) > 0 And InStr(report, "pack.zi_/secret.xlsm") > 0 Then
        Pass "実圧縮ZIPの入れ子と日本語名"
    Else
        Fail "実圧縮ZIPの入れ子と日本語名", report
    End If

    renamed = InspectAttachmentFile(renamedPath)
    If LineOf(renamed, "outerExtension: ") = ".zi_" And EntriesOf(report) = EntriesOf(renamed) And Not HasErrors(renamed) Then
        Pass "実圧縮ZIPの改名"
    Else
        Fail "実圧縮ZIPの改名", renamed
    End If

End Sub


'========================================================
' 判定
'========================================================

Private Sub CheckContainerExcluded(ByVal testRoot As String)

    Dim dirPath As String
    Dim innerPath As String
    Dim outerPath As String
    Dim mixedPath As String
    Dim innerBytes() As Byte
    Dim report As String

    dirPath = testRoot & "\fixtures"
    innerPath = dirPath & "\pdf-inner.zip"
    outerPath = dirPath & "\pdf-outer.zip"
    mixedPath = dirPath & "\mixed-outer.zip"

    ClearParts
    AddTextPart "a.pdf", "%PDF-a", False, False
    AddTextPart "b.pdf", "%PDF-b", False, False
    SaveParts innerPath
    innerBytes = ReadFileBytes(innerPath)

    ClearParts
    AddRawPart "uchiwake.zip", innerBytes, False, False
    AddTextPart "cover.pdf", "%PDF-c", False, False
    SaveParts outerPath

    report = InspectAttachmentFile(outerPath)
    If Not HasErrors(report) _
        And LineOf(report, "extensions: ") = ".pdf" _
        And InStr(report, "uchiwake.zip" & vbTab & ".zip" & vbTab & "container") > 0 _
        And InStr(report, "uchiwake.zip/a.pdf") > 0 _
        And InStr(report, "cover.pdf") > 0 Then
        Pass "器のZIPは拡張子から除外"
    Else
        Fail "器のZIPは拡張子から除外", report
    End If

    ClearParts
    AddTextPart "a.pdf", "%PDF-a", False, False
    AddTextPart "memo.txt", "memo", False, False
    SaveParts innerPath
    innerBytes = ReadFileBytes(innerPath)

    ClearParts
    AddRawPart "uchiwake.zip", innerBytes, False, False
    SaveParts mixedPath

    report = InspectAttachmentFile(mixedPath)
    If Not HasErrors(report) And LineOf(report, "extensions: ") = ".pdf, .txt" Then
        Pass "器の中のPDF以外は残る"
    Else
        Fail "器の中のPDF以外は残る", report
    End If

End Sub

Private Sub ExpectExts( _
    ByVal report As String, _
    ByVal caseName As String, _
    ParamArray expected() As Variant)

    Dim actual As String
    Dim parts() As String
    Dim i As Long
    Dim found As Boolean
    Dim j As Long

    actual = LineOf(report, "extensions: ")
    If Len(actual) = 0 Then
        Fail caseName, report
        Exit Sub
    End If

    parts = Split(actual, ", ")
    If UBound(parts) - LBound(parts) + 1 <> UBound(expected) - LBound(expected) + 1 Then
        Fail caseName, actual
        Exit Sub
    End If

    For i = LBound(expected) To UBound(expected)
        found = False
        For j = LBound(parts) To UBound(parts)
            If StrComp(parts(j), CStr(expected(i)), vbTextCompare) = 0 Then found = True
        Next j
        If Not found Then
            Fail caseName, "missing " & CStr(expected(i)) & " in " & actual
            Exit Sub
        End If
    Next i

    Pass caseName

End Sub


Private Function EntriesOf(ByVal report As String) As String

    Dim p As Long
    Dim q As Long

    p = InStr(report, "entries:" & vbCrLf)
    If p = 0 Then Exit Function
    p = p + Len("entries:" & vbCrLf)
    q = InStr(p, report, "errors:" & vbCrLf)
    If q = 0 Then
        EntriesOf = Mid$(report, p)
    Else
        EntriesOf = Mid$(report, p, q - p)
    End If

End Function


Private Function LineOf(ByVal report As String, ByVal prefix As String) As String

    Dim rows() As String
    Dim i As Long

    rows = Split(report, vbCrLf)
    For i = LBound(rows) To UBound(rows)
        If Left$(rows(i), Len(prefix)) = prefix Then
            LineOf = Mid$(rows(i), Len(prefix) + 1)
            Exit Function
        End If
    Next i

End Function


Private Function HasErrors(ByVal report As String) As Boolean

    HasErrors = (InStr(report, "errors:") > 0)

End Function


Private Function CountText(ByVal text As String, ByVal part As String) As Long

    Dim n As Long
    Dim p As Long

    p = 1
    Do
        p = InStr(p, text, part, vbTextCompare)
        If p = 0 Then Exit Do
        n = n + 1
        p = p + Len(part)
    Loop
    CountText = n

End Function


Private Sub Pass(ByVal caseName As String)

    mPass = mPass + 1
    mLog = mLog & "PASS " & caseName & vbCrLf

End Sub


Private Sub Fail(ByVal caseName As String, ByVal detail As String)

    mFail = mFail + 1
    mLog = mLog & "FAIL " & caseName & vbCrLf & detail & vbCrLf

End Sub


'========================================================
' テスト用 ZIP
'========================================================

Private Sub ClearParts()

    mPartCount = 0
    Erase mMeta
    BufInit mLocal

End Sub


Private Sub AddTextPart( _
    ByVal name As String, _
    ByVal text As String, _
    ByVal forceUtf8 As Boolean, _
    ByVal encrypted As Boolean)

    Dim raw() As Byte

    raw = AnsiBytes(text)
    AddRawPart name, raw, forceUtf8, encrypted

End Sub


Private Sub AddRawPart( _
    ByVal name As String, _
    ByRef payload() As Byte, _
    ByVal forceUtf8 As Boolean, _
    ByVal encrypted As Boolean)

    Dim nameBytes() As Byte
    Dim flags As Long
    Dim crc As Long
    Dim length As Long

    If forceUtf8 Then
        nameBytes = Utf8Bytes(name)
        flags = &H800
    Else
        nameBytes = AnsiBytes(name)
        flags = 0
    End If
    If encrypted Then flags = flags Or &H1

    length = ByteCount(payload)
    crc = Crc32(payload, length)

    mPartCount = mPartCount + 1
    ReDim Preserve mMeta(1 To mPartCount)
    mMeta(mPartCount).Name = name
    mMeta(mPartCount).ForceUtf8 = forceUtf8
    mMeta(mPartCount).Encrypted = encrypted
    mMeta(mPartCount).Crc = crc
    mMeta(mPartCount).Length = length
    mMeta(mPartCount).Offset = mLocal.Length

    BufAddU32 mLocal, &H4034B50
    BufAddU16 mLocal, 20
    BufAddU16 mLocal, flags
    BufAddU16 mLocal, 0
    BufAddU16 mLocal, 0
    BufAddU16 mLocal, 0
    BufAddU32 mLocal, crc
    BufAddU32 mLocal, length
    BufAddU32 mLocal, length
    BufAddU16 mLocal, ByteCount(nameBytes)
    BufAddU16 mLocal, 0
    BufAddBytes mLocal, nameBytes
    If length > 0 Then BufAddBytes mLocal, payload

End Sub


Private Sub SaveParts(ByVal filePath As String)

    Dim i As Long
    Dim nameBytes() As Byte
    Dim flags As Long
    Dim cdOffset As Long
    Dim cdSize As Long

    If mPartCount = 0 Then Exit Sub

    cdOffset = mLocal.Length

    For i = 1 To mPartCount
        If mMeta(i).ForceUtf8 Then
            nameBytes = Utf8Bytes(mMeta(i).Name)
            flags = &H800
        Else
            nameBytes = AnsiBytes(mMeta(i).Name)
            flags = 0
        End If
        If mMeta(i).Encrypted Then flags = flags Or &H1

        BufAddU32 mLocal, &H2014B50
        BufAddU16 mLocal, 20
        BufAddU16 mLocal, 20
        BufAddU16 mLocal, flags
        BufAddU16 mLocal, 0
        BufAddU16 mLocal, 0
        BufAddU16 mLocal, 0
        BufAddU32 mLocal, mMeta(i).Crc
        BufAddU32 mLocal, mMeta(i).Length
        BufAddU32 mLocal, mMeta(i).Length
        BufAddU16 mLocal, ByteCount(nameBytes)
        BufAddU16 mLocal, 0
        BufAddU16 mLocal, 0
        BufAddU16 mLocal, 0
        BufAddU16 mLocal, 0
        BufAddU32 mLocal, 0
        BufAddU32 mLocal, mMeta(i).Offset
        BufAddBytes mLocal, nameBytes
    Next i

    cdSize = mLocal.Length - cdOffset

    BufAddU32 mLocal, &H6054B50
    BufAddU16 mLocal, 0
    BufAddU16 mLocal, 0
    BufAddU16 mLocal, mPartCount
    BufAddU16 mLocal, mPartCount
    BufAddU32 mLocal, cdSize
    BufAddU32 mLocal, cdOffset
    BufAddU16 mLocal, 0

    BufSave mLocal, filePath

End Sub


Private Function MakeDeflateScript() As String

    Dim memo As String

    memo = "([char]0x65E5 + [char]0x672C + [char]0x8A9E + [char]0x30E1 + [char]0x30E2 + '.txt')"

    MakeDeflateScript = _
        "param([string]$arg)" & vbCrLf & _
        "$ErrorActionPreference = 'Stop'" & vbCrLf & _
        "$pair = $arg.Split('|')" & vbCrLf & _
        "$src = $pair[0]" & vbCrLf & _
        "$dest = $pair[1]" & vbCrLf & _
        "$memo = Join-Path $src " & memo & vbCrLf & _
        "[System.IO.File]::WriteAllText($memo, 'jp')" & vbCrLf & _
        "Add-Type -AssemblyName System.IO.Compression.FileSystem" & vbCrLf & _
        "if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }" & vbCrLf & _
        "[System.IO.Compression.ZipFile]::CreateFromDirectory($src, $dest)" & vbCrLf

End Function


Private Function VerifyScript() As String

    VerifyScript = _
        "param([string]$arg)" & vbCrLf & _
        "$ErrorActionPreference = 'Stop'" & vbCrLf & _
        "Add-Type -AssemblyName System.IO.Compression" & vbCrLf & _
        "Add-Type -AssemblyName System.IO.Compression.FileSystem" & vbCrLf & _
        "$zip = [System.IO.Compression.ZipFile]::OpenRead($arg)" & vbCrLf & _
        "try {" & vbCrLf & _
        "  foreach ($entry in $zip.Entries) {" & vbCrLf & _
        "    if ($entry.FullName.EndsWith('/')) { continue }" & vbCrLf & _
        "    $stream = $entry.Open()" & vbCrLf & _
        "    try {" & vbCrLf & _
        "      $buffer = New-Object byte[] 8192" & vbCrLf & _
        "      while ($true) {" & vbCrLf & _
        "        $read = $stream.Read($buffer, 0, $buffer.Length)" & vbCrLf & _
        "        if ($read -le 0) { break }" & vbCrLf & _
        "      }" & vbCrLf & _
        "    } finally { $stream.Dispose() }" & vbCrLf & _
        "  }" & vbCrLf & _
        "} finally { $zip.Dispose() }" & vbCrLf

End Function


Private Function RunPowerShellFile( _
    ByVal ps1Path As String, _
    ByVal script As String, _
    ByVal arg As String) As Long

    Dim sh As Object

    WriteAsciiFile ps1Path, script
    Set sh = CreateObject("WScript.Shell")
    RunPowerShellFile = sh.Run( _
        "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & _
        ps1Path & """ """ & arg & """", _
        0, True)

End Function


'========================================================
' バイト列
'========================================================

Private Function JpMemoName() As String

    ' 日本語メモ.txt
    JpMemoName = ChrW(&H65E5) & ChrW(&H672C) & ChrW(&H8A9E&) & _
        ChrW(&H30E1) & ChrW(&H30E2) & ".txt"

End Function


Private Function JpEstimateName() As String

    ' 見積.txt
    JpEstimateName = ChrW(&H898B&) & ChrW(&H7A4D) & ".txt"

End Function


Private Function Utf8Bytes(ByVal text As String) As Byte()

    Dim stm As Object
    Dim raw As Variant

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.Position = 0
    stm.Type = 1
    stm.Position = 3
    raw = stm.Read
    stm.Close
    Utf8Bytes = raw

End Function


Private Function Utf8Text(ByRef data() As Byte) As String

    Dim stm As Object

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write data
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "utf-8"
    Utf8Text = stm.ReadText
    stm.Close

End Function


Private Function AnsiBytes(ByVal text As String) As Byte()

    AnsiBytes = StrConv(text, vbFromUnicode)

End Function


Private Function Crc32(ByRef data() As Byte, ByVal length As Long) As Long

    Dim crc As Long
    Dim i As Long
    Dim j As Long
    Dim startAt As Long

    crc = &HFFFFFFFF
    If length <= 0 Then
        Crc32 = 0
        Exit Function
    End If

    startAt = LBound(data)
    For i = startAt To startAt + length - 1
        crc = crc Xor data(i)
        For j = 1 To 8
            If (crc And 1) <> 0 Then
                crc = ((crc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                crc = crc Xor &HEDB88320
            Else
                crc = ((crc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
    Next i

    Crc32 = Not crc

End Function


Private Function U32Hex(ByVal value As Long) As String

    U32Hex = Right$("00000000" & Hex$(value), 8)

End Function


Private Sub BufInit(ByRef b As Buf)

    b.Length = 0
    ReDim b.Data(0 To 1023)

End Sub


Private Sub BufAddByte(ByRef b As Buf, ByVal value As Byte)

    If b.Length > UBound(b.Data) Then
        ReDim Preserve b.Data(0 To (UBound(b.Data) + 1) * 2 - 1)
    End If
    b.Data(b.Length) = value
    b.Length = b.Length + 1

End Sub


Private Sub BufAddU16(ByRef b As Buf, ByVal value As Long)

    BufAddByte b, CByte(value And &HFF)
    BufAddByte b, CByte((value \ &H100) And &HFF)

End Sub


Private Sub BufAddU32(ByRef b As Buf, ByVal value As Long)

    Dim hi As Long

    BufAddByte b, CByte(value And &HFF)
    BufAddByte b, CByte((value And &HFF00&) \ &H100)
    BufAddByte b, CByte((value And &HFF0000) \ &H10000)

    If (value And &H80000000) <> 0 Then
        hi = ((value And &H7FFFFFFF) \ &H1000000) Or &H80
    Else
        hi = value \ &H1000000
    End If
    BufAddByte b, CByte(hi And &HFF)

End Sub


Private Sub BufAddBytes(ByRef b As Buf, ByRef data() As Byte)

    Dim i As Long
    Dim n As Long

    n = ByteCount(data)
    For i = 0 To n - 1
        BufAddByte b, data(LBound(data) + i)
    Next i

End Sub


Private Sub BufSave(ByRef b As Buf, ByVal filePath As String)

    Dim exact() As Byte
    Dim i As Long

    If b.Length <= 0 Then Exit Sub
    ReDim exact(0 To b.Length - 1)
    For i = 0 To b.Length - 1
        exact(i) = b.Data(i)
    Next i
    WriteBytes filePath, exact

End Sub


Private Function ByteCount(ByRef data() As Byte) As Long

    On Error GoTo EmptyArr
    ByteCount = UBound(data) - LBound(data) + 1
    Exit Function

EmptyArr:
    ByteCount = 0

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

    If Fso().FileExists(filePath) Then Fso().DeleteFile filePath, True
    fn = FreeFile
    Open filePath For Binary Access Write As #fn
    Put #fn, 1, data
    Close #fn

End Sub


Private Sub WriteAsciiFile(ByVal filePath As String, ByVal text As String)

    Dim ts As Object

    Set ts = Fso().CreateTextFile(filePath, True, False)
    ts.Write text
    ts.Close

End Sub


Private Sub WriteUtf8File(ByVal filePath As String, ByVal text As String)

    Dim stm As Object

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile filePath, 2
    stm.Close

End Sub


Private Function Fso() As Object

    If mFso Is Nothing Then
        Set mFso = CreateObject("Scripting.FileSystemObject")
    End If
    Set Fso = mFso

End Function
