'========================================================
' frmSendCheck - 送信前チェック UserForm
'
' 入れ方:
'   空の UserForm を作り、名前を frmSendCheck にする。
'   フォーム上には部品を置かない。このコードを貼る。
'   起動時に VBA が画面を自動生成する。
'
' 基本サイズ:
'   幅 800pt
'   高さ 460pt
' 画面の作業領域に収まらないときは開かない。
' フッターの 1.5 倍は、収まるときだけ全体を拡大する。
'========================================================
Option Explicit

'========================================================
' 送信前チェック UserForm
'
' UserForm上には何も部品を配置しません。
' 起動時にVBAで画面を自動生成します。
'
' 想定：
'   フルHD 1920 x 1080
'   Windows表示倍率 150%
'
' UserFormサイズ：
'   幅 800pt
'   高さ 460pt
'========================================================

Private mButtonHandlers As Collection
Private mCardHandlers As Collection
Private mAttachHandlers As Collection
Private mScrollDrags As Collection
Private mExpandedCard As MSForms.Frame
Private mCategoryFilter As String
Private mAddressFilter As String

Private Type AttachItem
    Id As String
    ParentId As String
    Caption As String
    Status As String
    Kind As String
    Depth As Long
    Expandable As Boolean
    PasswordState As String
    Checked As Boolean
End Type

Private mAttachItems() As AttachItem
Private mAttachCount As Long
Private mCollapsedIds As Collection
Private mAttachExpanded As Boolean
Private mCheckedPeople As Collection

Private mScale As Single
Private FORM_W As Single
Private FORM_H As Single

Private MARGIN As Single
Private GAP As Single

Private ATTACH_H As Single
Private ATTACH_HEADER As Single
Private ATTACH_ROW As Single

' カードの間隔。狭い列は5件、列を広げたときは3件が収まる
Private CARD_GAP As Single
Private CARD_PAD As Single

' リストの右に、スクロールバー専用の溝を空ける
Private SCROLL_LANE As Single

Private Sub LoadMetrics()

    If mScale <= 0 Then mScale = 1
    FORM_W = 800 * mScale
    FORM_H = 460 * mScale
    MARGIN = 12 * mScale
    GAP = 10 * mScale
    ATTACH_H = 78 * mScale
    ATTACH_HEADER = 22 * mScale
    ATTACH_ROW = 22 * mScale
    CARD_GAP = 3 * mScale
    CARD_PAD = 3 * mScale
    SCROLL_LANE = 10 * mScale

End Sub


Private Function S(ByVal v As Single) As Single

    If mScale <= 0 Then mScale = 1
    S = v * mScale

End Function

Private Sub UserForm_Initialize()

    '--------------------------------------------
    ' UserForm基本設定
    '--------------------------------------------
    Me.Caption = "送信前チェック"
    mScale = SendCheckScale
    If mScale <= 0 Then mScale = 1
    LoadMetrics
    Me.Width = FORM_W
    Me.Height = FORM_H

    Me.StartUpPosition = 1       '中央表示

    Me.BackColor = RGB(244, 245, 247)

    Set mButtonHandlers = New Collection
    Set mCardHandlers = New Collection
    Set mAttachHandlers = New Collection
    Set mScrollDrags = New Collection
    Set mCollapsedIds = New Collection
    Set mCheckedPeople = New Collection

    '--------------------------------------------
    ' 画面を作成
    '--------------------------------------------
    CreateRecipientArea

    CreateAttachmentArea

    CreateFooter

    ' マウスホイールで宛先列をスクロールする
    RestoreChecks
    WheelScroll_Attach Me

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    WheelScroll_Detach

End Sub


Private Sub UserForm_Terminate()

    WheelScroll_Detach

End Sub


'========================================================
' 宛先3列
'========================================================
Private Sub CreateRecipientArea()

    Dim colWidth As Single
    Dim topPos As Single
    Dim areaHeight As Single
    Dim external As Variant
    Dim outside As Variant
    Dim internal As Variant
    Dim toCount As Long
    Dim ccCount As Long
    Dim bccCount As Long

    colWidth = S(246)

    external = Array( _
        Array("山田 太郎", "ABC株式会社", "営業本部第一営業部法人営業室", "課長", "taro.yamada.a1b@asahikasei.com", "To"), _
        Array("John Smith", "XYZ Corporation", "Global Sales Division Enterprise Accounts", "Manager", "john.smith.b2c@asahikasei.com", "Cc"))
    outside = Array( _
        Array("佐藤 一郎", "株式会社サンプル", "生産技術本部プロセス革新部設備技術室", "主任", "ichiro.sato.c3d@asahikasei.com", "To"), _
        Array("田中 美咲", "株式会社サンプル", "経営企画本部コーポレート統括部総務企画室", "係長", "misaki.tanaka.d4e@asahikasei.com", "Cc"), _
        Array("高橋 健", "株式会社サンプル", "財務経理本部グローバル経理部連結決算室", "一般", "ken.takahashi.e5f@asahikasei.com", "Bcc"))
    internal = Array( _
        Array("鈴木 花子", "株式会社サンプル", "情報システム本部業務基盤部システム運用室", "課長", "hanako.suzuki.f6g@asahikasei.com", "To"), _
        Array("高橋 翔", "株式会社サンプル", "研究開発本部素材科学部高分子研究室", "主任", "sho.takahashi.g7h@asahikasei.com", "Cc"), _
        Array("伊藤 健太", "株式会社サンプル", "製造本部生産管理部工程計画室", "係長", "kenta.ito.h8i@asahikasei.com", "Cc"), _
        Array("渡辺 美香", "株式会社サンプル", "品質保証本部製品安全部監査推進室", "一般", "mika.watanabe.i9j@asahikasei.com", "Bcc"), _
        Array("小林 直樹", "株式会社サンプル", "人事本部人材開発部研修企画室", "一般", "naoki.kobayashi.j0k@asahikasei.com", "Bcc"), _
        Array("中村 葵", "株式会社サンプル", "営業本部第二営業部広域法人営業室", "課長", "aoi.nakamura.k1l@asahikasei.com", "To"), _
        Array("加藤 蓮", "株式会社サンプル", "法務・コンプライアンス本部契約管理部", "主任", "ren.kato.l2m@asahikasei.com", "Cc"), _
        Array("吉田 さくら", "株式会社サンプル", "コミュニケーション本部広報部報道室", "一般", "sakura.yoshida.m3n@asahikasei.com", "Bcc"), _
        Array("山本 大輔", "株式会社サンプル", "調達本部原材料部購買企画室", "係長", "daisuke.yamamoto.n4o@asahikasei.com", "To"))

    CountAddressTypes external, toCount, ccCount, bccCount
    CountAddressTypes outside, toCount, ccCount, bccCount
    CountAddressTypes internal, toCount, ccCount, bccCount
    CreateRoleSummary toCount, ccCount, bccCount

    topPos = S(30)
    areaHeight = (Me.InsideHeight - MARGIN - S(28) - GAP - ATTACH_H) - GAP - topPos

    '--------------------------------------------
    ' 社外
    '--------------------------------------------
    CreateCategoryFrame _
        "fraExternal", _
        "社外", _
        "2", _
        MARGIN, _
        topPos, _
        colWidth, _
        areaHeight, _
        RGB(176, 64, 60), _
        external

    '--------------------------------------------
    ' 部外
    '--------------------------------------------
    CreateCategoryFrame _
        "fraOutside", _
        "部外", _
        "3", _
        MARGIN + colWidth + GAP, _
        topPos, _
        colWidth, _
        areaHeight, _
        RGB(156, 112, 20), _
        outside

    '--------------------------------------------
    ' 部内
    '--------------------------------------------
    CreateCategoryFrame _
        "fraInternal", _
        "部内", _
        "9", _
        MARGIN + (colWidth + GAP) * 2, _
        topPos, _
        colWidth, _
        areaHeight, _
        RGB(42, 96, 148), _
        internal

    PlaceInternalCheckButton

End Sub


Private Sub CountAddressTypes( _
    ByVal recipients As Variant, _
    ByRef toCount As Long, _
    ByRef ccCount As Long, _
    ByRef bccCount As Long)

    Dim i As Long
    Dim addressType As String

    For i = LBound(recipients) To UBound(recipients)
        addressType = CStr(recipients(i)(5))

        If StrComp(addressType, "To", vbTextCompare) = 0 Then
            toCount = toCount + 1
        ElseIf StrComp(addressType, "Cc", vbTextCompare) = 0 Then
            ccCount = ccCount + 1
        ElseIf StrComp(addressType, "Bcc", vbTextCompare) = 0 Then
            bccCount = bccCount + 1
        End If
    Next i

End Sub


Private Sub CreateRoleSummary( _
    ByVal toCount As Long, _
    ByVal ccCount As Long, _
    ByVal bccCount As Long)

    Dim lbl As MSForms.Label
    Dim texts(1 To 3) As String
    Dim keys(1 To 3) As String
    Dim widths(1 To 3) As Single
    Dim i As Long
    Dim leftPos As Single
    Dim textW As Single
    Dim visualGap As Single

    keys(1) = "To"
    keys(2) = "Cc"
    keys(3) = "Bcc"
    texts(1) = AddressTypeCaption(keys(1)) & " " & toCount
    texts(2) = AddressTypeCaption(keys(2)) & " " & ccCount
    texts(3) = AddressTypeCaption(keys(3)) & " " & bccCount
    widths(1) = S(52)
    widths(2) = S(52)
    widths(3) = S(60)

    leftPos = MARGIN

    For i = 1 To 3
        Set lbl = AddLabel( _
            "role" & keys(i), _
            texts(i), _
            leftPos, S(6), widths(i), S(20))

        With lbl
            .Tag = keys(i)
            .Font.Name = "Meiryo UI"
            .Font.Size = S(10)
            .Font.Bold = True
            .ForeColor = RGB(55, 55, 60)
            .BackStyle = fmBackStyleTransparent
            .ControlTipText = AddressTypeCaption(keys(i)) & " の人だけ表示"
        End With

        RegisterCardClick lbl
        textW = MeasureLabelWidth(lbl)
        lbl.Width = textW + S(2)
        lbl.Height = S(20)
        visualGap = (widths(i) - lbl.Width) + S(14)
        If visualGap < S(14) Then visualGap = S(14)
        leftPos = leftPos + lbl.Width + (visualGap * 2 / 3)
    Next i

    ' BCC の右の空白。クリックで表示を初期化する
    Set lbl = AddLabel( _
        "lblCheckRemain", _
        "", _
        Me.InsideWidth - MARGIN - S(20), S(6), S(20), S(20))

    With lbl
        .Font.Name = "Meiryo UI"
        .Font.Size = S(10)
        .Font.Bold = False
        .ForeColor = RGB(176, 64, 60)
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .TextAlign = fmTextAlignCenter
    End With

    Set lbl = AddLabel( _
        "lblHeaderBlank", _
        "", _
        leftPos, 0, Me.InsideWidth - leftPos, S(28))

    With lbl
        .BackStyle = fmBackStyleTransparent
        .ControlTipText = "表示を初期化"
    End With

    RegisterCardClick lbl

    Set lbl = AddLabel( _
        "lblHeaderBlankRight", _
        "", _
        Me.InsideWidth - S(8), 0, S(8), S(28))

    With lbl
        .BackStyle = fmBackStyleTransparent
        .ControlTipText = Me.Controls("lblHeaderBlank").ControlTipText
    End With

    RegisterCardClick lbl

    UpdateRoleSummaryStyle

End Sub


Private Function AddressTypeCaption(ByVal addressType As String) As String

    Select Case addressType
        Case "Cc"
            AddressTypeCaption = "CC"
        Case "Bcc"
            AddressTypeCaption = "BCC"
        Case Else
            AddressTypeCaption = addressType
    End Select

End Function


Public Sub ToggleAddressFilter(ByVal addressType As String)

    If mAddressFilter = addressType Then
        mAddressFilter = ""
    Else
        mAddressFilter = addressType
    End If

    ApplyAddressFilter

End Sub


Private Sub ApplyAddressFilter()

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim ctl As Object
    Dim shown As Long

    If Not mExpandedCard Is Nothing Then
        If mAddressFilter <> "" Then
            If StrComp(CStr(mExpandedCard.Tag), mAddressFilter, vbTextCompare) <> 0 Then
                ApplyCardLayout mExpandedCard, False
                Set mExpandedCard = Nothing
            End If
        End If
    End If

    names = Array("fraExternal", "fraOutside", "fraInternal")

    For i = LBound(names) To UBound(names)
        Set fra = Me.Controls(CStr(names(i)))
        Set list = fra.Controls(fra.Name & "_list")
        shown = 0

        For Each ctl In list.Controls
            If TypeName(ctl) = "Frame" Then
                If mAddressFilter = "" Or _
                    StrComp(CStr(ctl.Tag), mAddressFilter, vbTextCompare) = 0 Then
                    ctl.Visible = True
                    shown = shown + 1
                Else
                    ctl.Visible = False
                End If
            End If
        Next ctl

        fra.Controls(fra.Name & "_count").Caption = CStr(shown)
        ReflowCards list
        list.ScrollTop = 0
        SyncScrollThumb list
    Next i

    UpdateRoleSummaryStyle

End Sub


Private Sub UpdateRoleSummaryStyle()

    Dim keys As Variant
    Dim i As Long
    Dim lbl As MSForms.Label
    Dim selected As Boolean

    keys = Array("To", "Cc", "Bcc")

    For i = LBound(keys) To UBound(keys)
        Set lbl = Me.Controls("role" & CStr(keys(i)))
        selected = (mAddressFilter = CStr(keys(i)))

        If mAddressFilter = "" Or selected Then
            lbl.ForeColor = RGB(55, 55, 60)
        Else
            lbl.ForeColor = RGB(180, 180, 186)
        End If

        lbl.Font.Underline = selected

        If selected Then
            lbl.ControlTipText = "全員を表示"
        Else
            lbl.ControlTipText = AddressTypeCaption(CStr(keys(i))) & " の人だけ表示"
        End If
    Next i

End Sub


'========================================================
' カテゴリFrame
'========================================================
Private Sub CreateCategoryFrame( _
    ByVal frameName As String, _
    ByVal title As String, _
    ByVal countText As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal frameWidth As Single, _
    ByVal frameHeight As Single, _
    ByVal titleColor As Long, _
    ByVal recipients As Variant)

    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim lbl As MSForms.Label

    Dim i As Long
    Dim cardTop As Single
    Dim headerHeight As Single
    Dim person As Object
    Dim people() As MSForms.Frame
    Dim n As Long

    headerHeight = S(34)

    '--------------------------------------------
    ' 外枠。タイトルはここに置き、スクロールしない
    '--------------------------------------------
    Set fra = Me.Controls.Add( _
        "Forms.Frame.1", _
        frameName, _
        True)

    With fra
        .Left = leftPos
        .Top = topPos
        .Width = frameWidth
        .Height = frameHeight

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(226, 227, 230)
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
    End With

    '--------------------------------------------
    ' カテゴリタイトル
    '--------------------------------------------
    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        frameName & "_title", _
        True)

    With lbl
        .Left = S(10)
        .Top = S(8)
        .Width = S(155)
        .Height = S(22)

        .Caption = title

        .Font.Name = "Meiryo UI"
        .Font.Size = S(11)
        .Font.Bold = True

        .ForeColor = titleColor

        .BackStyle = fmBackStyleTransparent
    End With

    '人数
    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        frameName & "_count", _
        True)

    With lbl
        .Left = frameWidth - S(58)
        .Top = S(9)
        .Width = S(42)
        .Height = S(18)

        .Caption = countText

        .Font.Name = "Meiryo UI"
        .Font.Size = S(8)
        .Font.Bold = False

        .ForeColor = RGB(140, 140, 146)

        .TextAlign = fmTextAlignRight

        .BackStyle = fmBackStyleTransparent
    End With

    '--------------------------------------------
    ' 宛先だけをスクロールする領域
    '--------------------------------------------
    Set list = fra.Controls.Add( _
        "Forms.Frame.1", _
        frameName & "_list", _
        True)

    With list
        .Left = S(1)
        .Top = headerHeight
        .Width = frameWidth - S(2) - SCROLL_LANE
        .Height = frameHeight - headerHeight - S(1)

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
        .ScrollTop = 0
        .ScrollLeft = 0
    End With

    '--------------------------------------------
    ' 宛先カード
    '--------------------------------------------
    cardTop = CARD_PAD

    For i = LBound(recipients) To UBound(recipients)

        CreatePersonCard _
            list, _
            i + 1, _
            cardTop, _
            recipients(i)(0), _
            recipients(i)(1), _
            recipients(i)(2), _
            recipients(i)(3), _
            recipients(i)(4), _
            recipients(i)(5), _
            titleColor

    Next i

    ' 列の高さに合わせて、5件分のカード高さにする
    n = 0
    For Each person In list.Controls
        If TypeName(person) = "Frame" Then
            n = n + 1
            ReDim Preserve people(1 To n)
            Set people(n) = person
        End If
    Next

    For i = 1 To n
        ApplyCardLayout people(i), False
    Next i

    RegisterCardClick fra
    RegisterCardClick fra.Controls(frameName & "_title")
    RegisterCardClick fra.Controls(frameName & "_count")
    fra.Controls(frameName & "_title").ControlTipText = "この列だけ表示"

End Sub


'========================================================
' 人物カード
'========================================================
Private Sub CreatePersonCard( _
    ByVal parentFrame As MSForms.Frame, _
    ByVal cardNo As Long, _
    ByVal topPos As Single, _
    ByVal personName As String, _
    ByVal companyName As String, _
    ByVal departmentName As String, _
    ByVal jobTitle As String, _
    ByVal emailAddress As String, _
    ByVal addressType As String, _
    ByVal categoryColor As Long)

    Dim card As MSForms.Frame
    Dim photo As MSForms.Label
    Dim lbl As MSForms.Label

    Dim cardWidth As Single

    cardWidth = parentFrame.Width - S(16)

    '--------------------------------------------
    ' カード本体
    '--------------------------------------------
    Set card = parentFrame.Controls.Add( _
        "Forms.Frame.1", _
        "card_" & cardNo, _
        True)

    With card
        .Left = S(8)
        .Top = topPos
        .Width = cardWidth
        .Height = S(92)

        .Caption = ""
        .Tag = addressType

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(226, 227, 230)
        .SpecialEffect = fmSpecialEffectFlat
    End With

    '--------------------------------------------
    ' 列の上からの番号
    '--------------------------------------------
    Set lbl = card.Controls.Add( _
        "Forms.Label.1", _
        "no_" & cardNo, _
        True)

    With lbl
        .Left = 0
        .Top = 0
        .Width = S(18)
        .Height = S(16)

        .Caption = CStr(cardNo)

        .Font.Name = "Meiryo UI"
        .Font.Size = S(8)
        .Font.Bold = True

        .ForeColor = categoryColor
        .BackStyle = fmBackStyleTransparent

        .TextAlign = fmTextAlignLeft

        .BorderStyle = fmBorderStyleNone
    End With

    '--------------------------------------------
    ' プロフィール写真
    '
    ' 現在は仮表示。
    ' 後でMicrosoft 365のプロフィール画像に置換する。
    '--------------------------------------------
    Set photo = card.Controls.Add( _
        "Forms.Label.1", _
        "photo_" & cardNo, _
        True)

    With photo
        .Left = S(8)
        .Top = S(25)
        .Width = S(42)
        .Height = S(42)

        .Caption = GetInitials(personName)
        .Tag = .Caption

        .Font.Name = "Meiryo UI"
        .Font.Size = S(12)
        .Font.Bold = True

        .ForeColor = RGB(255, 255, 255)
        .BackColor = categoryColor

        .TextAlign = fmTextAlignCenter

        .BorderStyle = fmBorderStyleNone
    End With

    ' 確認チェック。画像の右下に、白地の小さい印を重ねる
    Set lbl = card.Controls.Add( _
        "Forms.Label.1", _
        "phshd_" & cardNo, _
        True)

    With lbl
        .Caption = ""
        .Visible = False
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
    End With

    Set lbl = card.Controls.Add( _
        "Forms.Label.1", _
        "phchk_" & cardNo, _
        True)

    With lbl
        .Caption = ChrW(&HEC61)
        .Visible = False
        .Font.Name = "Segoe MDL2 Assets"
        .Font.Bold = False
        .Font.Size = S(14)
        .ForeColor = RGB(255, 255, 255)
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleNone
        .WordWrap = False
    End With

    '--------------------------------------------
    ' 氏名
    '--------------------------------------------
    Set lbl = card.Controls.Add( _
        "Forms.Label.1", _
        "name_" & cardNo, _
        True)

    With lbl
        .Left = S(60)
        .Top = S(6)
        .Width = cardWidth - S(158)
        .Height = S(16)

        .Caption = personName

        .Font.Name = "Meiryo UI"
        .Font.Size = S(9)
        .Font.Bold = True

        .ForeColor = RGB(28, 28, 30)

        .BackStyle = fmBackStyleTransparent
    End With

    '--------------------------------------------
    ' 会社名
    ' 社外だけ赤の太字
    '--------------------------------------------
    If categoryColor = RGB(176, 64, 60) Then
        AddCardDetail card, "company_" & cardNo, companyName, S(60), S(22), cardWidth - S(68), categoryColor, True
    Else
        AddCardDetail card, "company_" & cardNo, companyName, S(60), S(22), cardWidth - S(68), RGB(70, 70, 76)
    End If
    card.Controls("company_" & cardNo).Tag = companyName

    '--------------------------------------------
    ' 部署名
    '--------------------------------------------
    AddCardDetail card, "dept_" & cardNo, departmentName, S(60), S(36), cardWidth - S(68), RGB(110, 110, 116)
    card.Controls("dept_" & cardNo).Tag = departmentName

    '--------------------------------------------
    ' 役職
    '--------------------------------------------
    AddCardDetail card, "title_" & cardNo, jobTitle, S(60), S(50), cardWidth - S(68), RGB(110, 110, 116)

    '--------------------------------------------
    ' メールアドレス
    '--------------------------------------------
    AddCardDetail card, "mail_" & cardNo, emailAddress, S(60), S(64), cardWidth - S(68), RGB(140, 140, 146)

    '--------------------------------------------
    ' To / CC / BCC
    ' 設定されているものだけ強調し、残りは薄くする
    '--------------------------------------------
    CreateAddressTypeMarks card, cardNo, cardWidth, addressType, categoryColor

    RegisterCardClicks card

End Sub


Private Sub AddCardDetail( _
    ByVal card As MSForms.Frame, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal width As Single, _
    ByVal foreColor As Long, _
    Optional ByVal isBold As Boolean = False)

    Dim lbl As MSForms.Label

    Set lbl = card.Controls.Add( _
        "Forms.Label.1", _
        controlName, _
        True)

    With lbl
        .Left = leftPos
        .Top = topPos
        .Width = width
        .Height = S(14)

        .Caption = caption

        .Font.Name = "Meiryo UI"
        .Font.Size = S(8)
        .Font.Bold = isBold

        .ForeColor = foreColor

        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
    End With

End Sub


Private Sub RegisterCardClicks(ByVal card As MSForms.Frame)

    Dim ctl As Object

    RegisterCardClick card

    For Each ctl In card.Controls
        If TypeName(ctl) = "Label" Then
            RegisterCardClick ctl
        End If
    Next ctl

End Sub


Private Sub RegisterCardClick(ByVal ctl As Object)

    Dim handler As CCardClick

    Set handler = New CCardClick
    Set handler.ParentForm = Me

    If TypeName(ctl) = "Frame" Then
        Set handler.CardFrame = ctl
    Else
        Set handler.CardLabel = ctl
    End If

    mCardHandlers.Add handler

End Sub


Public Sub ToggleCategoryFilter(ByVal fra As MSForms.Frame)

    If mCategoryFilter = fra.Name Then
        mCategoryFilter = ""
    Else
        mCategoryFilter = fra.Name
    End If

    ApplyCategoryFilter

End Sub


Private Sub ApplyCategoryFilter()

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim colWidth As Single
    Dim fullWidth As Single

    colWidth = S(246)
    fullWidth = colWidth * 3 + GAP * 2
    names = Array("fraExternal", "fraOutside", "fraInternal")

    If Not mExpandedCard Is Nothing Then
        If mCategoryFilter <> "" Then
            If mExpandedCard.Parent.Parent.Name <> mCategoryFilter Then
                ApplyCardLayout mExpandedCard, False
                Set mExpandedCard = Nothing
            End If
        End If
    End If

    For i = LBound(names) To UBound(names)

        Set fra = Me.Controls(names(i))

        If mCategoryFilter = "" Or fra.Name = mCategoryFilter Then
            fra.Visible = True

            If mCategoryFilter = "" Then
                fra.Left = MARGIN + i * (colWidth + GAP)
                SizeCategoryColumn fra, colWidth
                fra.Controls(fra.Name & "_title").ControlTipText = "この列だけ表示"
            Else
                fra.Left = MARGIN
                SizeCategoryColumn fra, fullWidth
                fra.Controls(fra.Name & "_title").ControlTipText = "3列に戻す"
            End If
        Else
            fra.Visible = False
        End If

    Next i

End Sub


Private Sub SizeCategoryColumn( _
    ByVal fra As MSForms.Frame, _
    ByVal colWidth As Single)

    Dim list As MSForms.Frame
    Dim ctl As Object
    Dim cards() As MSForms.Frame
    Dim n As Long
    Dim i As Long

    fra.Width = colWidth
    fra.Controls(fra.Name & "_count").Left = colWidth - S(58)

    Set list = fra.Controls(fra.Name & "_list")
    list.Width = colWidth - S(2) - SCROLL_LANE

    n = 0
    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then n = n + 1
    Next ctl

    If n = 0 Then Exit Sub

    ReDim cards(1 To n)
    i = 0

    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then
            i = i + 1
            Set cards(i) = ctl
        End If
    Next ctl

    For i = 1 To n
        If list.Width > S(400) Then
            cards(i).Width = (list.Width - S(16) - GAP) / 2
            ApplyCardLayout cards(i), True
        Else
            cards(i).Width = list.Width - S(16)
            ApplyCardLayout cards(i), _
                (Not mExpandedCard Is Nothing) And (mExpandedCard Is cards(i))
        End If
    Next i

End Sub


Private Function CardIsWide(ByVal card As MSForms.Frame) As Boolean

    ' カード自体ではなく、列を広げたかどうかを見る
    CardIsWide = (card.Parent.Width > S(400))

End Function


Private Function CardSlotCount(ByVal wide As Boolean) As Long

    If wide Then
        CardSlotCount = 3
    Else
        CardSlotCount = 5
    End If

End Function


Private Function ListViewHeight(ByVal list As MSForms.Frame) As Single

    On Error Resume Next
    ListViewHeight = list.InsideHeight
    On Error GoTo 0
    If ListViewHeight < S(20) Then ListViewHeight = list.Height

End Function


' 指定件数ぴったり収まるカード高さ
Private Function FittedCardHeight( _
    ByVal list As MSForms.Frame, _
    ByVal slots As Long) As Single

    Dim view As Single

    If slots < 1 Then slots = 1
    view = ListViewHeight(list)
    FittedCardHeight = (view - CARD_PAD * 2 - CARD_GAP * (slots - 1)) / slots
    If FittedCardHeight < S(34) Then FittedCardHeight = S(34)

End Function


Private Sub RepositionCardChrome( _
    ByVal card As MSForms.Frame, _
    ByVal cardNo As String)

    Dim leftPos As Single
    Dim kinds As Variant
    Dim i As Long
    Dim nameWidth As Single
    Dim photo As MSForms.Label
    Dim nameLabel As MSForms.Label
    Dim wide As Boolean

    Dim photoSize As Single

    wide = CardIsWide(card)

    Set photo = card.Controls("photo_" & cardNo)
    Set nameLabel = card.Controls("name_" & cardNo)

    If wide Then
        photoSize = card.Height - S(18)
        If photoSize > S(48) Then photoSize = S(48)
        If photoSize < S(36) Then photoSize = S(36)
        photo.Left = S(18)
        photo.Font.Size = S(14)
        nameLabel.Top = S(4)
        nameLabel.Height = S(14)
        nameLabel.Font.Size = S(10)
    Else
        photoSize = card.Height - S(14)
        If photoSize > S(28) Then photoSize = S(28)
        If photoSize < S(18) Then photoSize = S(18)
        photo.Left = S(16)
        photo.Font.Size = S(9)
        nameLabel.Top = S(1)
        nameLabel.Height = S(12)
        nameLabel.Font.Size = S(9)
    End If

    photo.Width = photoSize
    photo.Height = photoSize
    If photo.Height > card.Height - S(4) Then
        photo.Height = card.Height - S(4)
        photo.Width = photo.Height
    End If
    If photo.Height < S(16) Then
        photo.Height = S(16)
        photo.Width = S(16)
    End If
    photo.Top = (card.Height - photo.Height) / 2
    If photo.Top < S(2) Then photo.Top = S(2)

    nameLabel.Left = photo.Left + photo.Width + S(6)
    nameWidth = card.Width - S(94) - nameLabel.Left - S(4)
    If nameWidth < S(36) Then nameWidth = S(36)
    nameLabel.Width = nameWidth

    ' 番号は写真の左、白地の上に色文字で置く
    With card.Controls("no_" & cardNo)
        .Left = S(1)
        .Top = S(2)
        .Width = S(14)
        .Height = S(12)
        .BackStyle = fmBackStyleTransparent
        .ForeColor = photo.BackColor
        .Font.Bold = True
        .Font.Size = S(8)
        .TextAlign = fmTextAlignLeft
    End With

    PlacePersonCheck card

    kinds = Array("To", "Cc", "Bcc")
    leftPos = card.Width - S(94)

    For i = LBound(kinds) To UBound(kinds)
        With card.Controls("type_" & cardNo & "_" & kinds(i))
            .Left = leftPos
            .Top = nameLabel.Top
        End With
        leftPos = leftPos + S(30)
    Next i

End Sub


Public Sub ToggleCard(ByVal card As MSForms.Frame)

    ' 列を広げた状態は、最初から全カードを展開している
    If CardIsWide(card) Then Exit Sub

    If Not mExpandedCard Is Nothing Then
        If mExpandedCard Is card Then
            ApplyCardLayout card, False
            Set mExpandedCard = Nothing
            Exit Sub
        End If

        ApplyCardLayout mExpandedCard, False
    End If

    ApplyCardLayout card, True
    Set mExpandedCard = card
    ScrollCardIntoView card

End Sub


Private Sub ApplyCardLayout( _
    ByVal card As MSForms.Frame, _
    ByVal expanded As Boolean)

    Dim cardNo As String
    Dim textWidth As Single
    Dim textLeft As Single
    Dim topPos As Single
    Dim minHeight As Single
    Dim fields(1 To 4) As String
    Dim i As Long
    Dim lbl As MSForms.Label
    Dim wide As Boolean
    Dim nameLabel As MSForms.Label
    Dim photo As MSForms.Label

    cardNo = Mid(card.Name, 6)
    wide = CardIsWide(card)
    minHeight = FittedCardHeight(card.Parent, CardSlotCount(wide))
    card.Height = minHeight
    RepositionCardChrome card, cardNo

    Set nameLabel = card.Controls("name_" & cardNo)
    textLeft = nameLabel.Left

    fields(1) = "company_" & cardNo
    fields(2) = "dept_" & cardNo
    fields(3) = "title_" & cardNo
    fields(4) = "mail_" & cardNo

    If expanded Then
        card.Controls("company_" & cardNo).Caption = CStr(card.Controls("company_" & cardNo).Tag)
        textWidth = card.Width - textLeft - S(12)
        If textWidth < S(40) Then textWidth = S(40)
        topPos = nameLabel.Top + nameLabel.Height + S(2)

        For i = 1 To 4
            Set lbl = card.Controls(fields(i))
            lbl.Visible = True
            lbl.Left = textLeft
            lbl.Top = topPos
            lbl.WordWrap = True
            lbl.AutoSize = False
            lbl.Width = textWidth
            lbl.Height = WrappedLineCount(lbl, textWidth) * S(14)
            topPos = lbl.Top + lbl.Height + S(1)
        Next i

        card.Height = topPos + S(4)
        If card.Height < minHeight Then card.Height = minHeight

        ' 展開後のカード中央より少し上。上に寄りすぎないように下げる
        If wide Then
            Set photo = card.Controls("photo_" & cardNo)
            photo.Top = (card.Height - photo.Height) / 2 - S(6)
            If photo.Top < S(12) Then photo.Top = S(12)
            If photo.Top + photo.Height > card.Height - S(8) Then
                photo.Top = card.Height - photo.Height - S(8)
            End If
            PlacePersonCheck card
        End If
    Else
        ' 初期表示は「会社名/部署名」を1行。見切れはそのまま
        textWidth = card.Width - textLeft - S(8)
        If textWidth < S(40) Then textWidth = S(40)

        Set lbl = card.Controls("company_" & cardNo)
        lbl.Caption = CStr(lbl.Tag) & "/" & CStr(card.Controls("dept_" & cardNo).Tag)
        lbl.Visible = True
        lbl.WordWrap = False
        lbl.AutoSize = False
        lbl.Left = textLeft
        lbl.Top = nameLabel.Top + nameLabel.Height + S(1)
        lbl.Width = textWidth
        lbl.Height = S(12)

        card.Controls("dept_" & cardNo).Visible = False
        card.Controls("title_" & cardNo).Visible = False
        card.Controls("mail_" & cardNo).Visible = False
    End If

    ReflowCards card.Parent

End Sub


Private Function WrappedLineCount( _
    ByVal lbl As MSForms.Label, _
    ByVal maxWidth As Single) As Long

    Dim fullWidth As Single

    lbl.WordWrap = False
    lbl.AutoSize = True
    fullWidth = lbl.Width
    lbl.AutoSize = False
    lbl.WordWrap = True
    lbl.Width = maxWidth

    If maxWidth <= 0 Then
        WrappedLineCount = 1
    Else
        WrappedLineCount = -Int(-fullWidth / maxWidth)
    End If

    If WrappedLineCount < 1 Then WrappedLineCount = 1

End Function


Private Sub ReflowCards(ByVal list As MSForms.Frame)

    Dim ctl As Object
    Dim cards() As MSForms.Frame
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As MSForms.Frame
    Dim topPos As Single
    Dim visibleCards() As MSForms.Frame
    Dim v As Long
    Dim rowH As Single
    Dim left0 As Single

    n = 0
    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then
            n = n + 1
        End If
    Next ctl

    If n = 0 Then Exit Sub

    ReDim cards(1 To n)
    i = 0
    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then
            i = i + 1
            Set cards(i) = ctl
        End If
    Next ctl

    For i = 1 To n - 1
        For j = i + 1 To n
            If CLng(Mid(cards(j).Name, 6)) < CLng(Mid(cards(i).Name, 6)) Then
                Set tmp = cards(i)
                Set cards(i) = cards(j)
                Set cards(j) = tmp
            End If
        Next j
    Next i

    v = 0
    For i = 1 To n
        If cards(i).Visible Then
            v = v + 1
            ReDim Preserve visibleCards(1 To v)
            Set visibleCards(v) = cards(i)
        End If
    Next i

    left0 = S(8)
    topPos = CARD_PAD

    If list.Width > S(400) And v > 0 Then
        i = 1
        Do While i <= v
            visibleCards(i).Left = left0
            visibleCards(i).Top = topPos
            rowH = visibleCards(i).Height
            If i < v Then
                visibleCards(i + 1).Left = left0 + visibleCards(i).Width + GAP
                visibleCards(i + 1).Top = topPos
                If visibleCards(i + 1).Height > rowH Then
                    rowH = visibleCards(i + 1).Height
                End If
                i = i + 2
            Else
                i = i + 1
            End If
            topPos = topPos + rowH + CARD_GAP
        Loop
    Else
        For i = 1 To v
            visibleCards(i).Left = left0
            visibleCards(i).Top = topPos
            topPos = topPos + visibleCards(i).Height + CARD_GAP
        Next i
    End If

    If v > 0 And topPos > CARD_PAD Then topPos = topPos - CARD_GAP
    list.ScrollHeight = topPos + CARD_PAD
    SyncScrollThumb list

End Sub


Private Sub ScrollCardIntoView(ByVal card As MSForms.Frame)

    Dim list As MSForms.Frame
    Dim viewHeight As Single
    Dim newTop As Single

    Set list = card.Parent
    viewHeight = list.InsideHeight
    If viewHeight <= 0 Then viewHeight = list.Height

    If card.Height >= viewHeight Then
        list.ScrollTop = card.Top
        SyncScrollThumb list
        Exit Sub
    End If

    If card.Top < list.ScrollTop Then
        list.ScrollTop = card.Top
    ElseIf card.Top + card.Height > list.ScrollTop + viewHeight Then
        newTop = card.Top + card.Height - viewHeight
        If newTop < 0 Then newTop = 0
        list.ScrollTop = newTop
    End If

    SyncScrollThumb list

End Sub


Private Sub CreateAddressTypeMarks( _
    ByVal card As MSForms.Frame, _
    ByVal cardNo As Long, _
    ByVal cardWidth As Single, _
    ByVal addressType As String, _
    ByVal categoryColor As Long)

    Dim kinds(1 To 3) As String
    Dim i As Long
    Dim lbl As MSForms.Label
    Dim markWidth As Single
    Dim gap As Single
    Dim leftPos As Single
    Dim selected As Boolean

    kinds(1) = "To"
    kinds(2) = "Cc"
    kinds(3) = "Bcc"

    markWidth = S(28)
    gap = S(2)
    leftPos = cardWidth - (markWidth * 3 + gap * 2) - S(6)

    For i = 1 To 3

        selected = (StrComp(addressType, kinds(i), vbTextCompare) = 0)

        Set lbl = card.Controls.Add( _
            "Forms.Label.1", _
            "type_" & cardNo & "_" & kinds(i), _
            True)

        With lbl
            .Left = leftPos
            .Top = S(5)
            .Width = markWidth
            .Height = S(16)

            .Caption = AddressTypeCaption(kinds(i))

            .Font.Name = "Meiryo UI"
            .Font.Bold = selected

            If selected Then
                .Font.Size = S(9)
            Else
                .Font.Size = S(7)
            End If

            .TextAlign = fmTextAlignCenter

            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone

            If selected Then
                .ForeColor = categoryColor
            Else
                .ForeColor = RGB(196, 196, 200)
            End If
        End With

        leftPos = leftPos + markWidth + gap

    Next i

End Sub


'========================================================
' 添付ファイル
' ダミー。ZIP の行を押すと、その中身だけ開く。
'========================================================
Private Sub CreateAttachmentArea()

    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim lbl As MSForms.Label

    LoadDummyAttachments

    Set fra = Me.Controls.Add( _
        "Forms.Frame.1", _
        "fraAttachments", _
        True)

    With fra
        .Left = MARGIN
        .Top = Me.InsideHeight - MARGIN - S(28) - GAP - ATTACH_H
        .Width = S(246) * 3 + GAP * 2
        .Height = ATTACH_H

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(226, 227, 230)
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
    End With

    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        "lblAttachmentTitle", _
        True)

    With lbl
        .Left = S(10)
        .Top = S(2)
        .Width = S(220)
        .Height = S(18)

        .Caption = ChrW(&H25B8) & "  添付ファイル"
        .ControlTipText = "クリックで一覧を上まで広げる"

        .Font.Name = "Meiryo UI"
        .Font.Size = S(10)
        .Font.Bold = True

        .ForeColor = RGB(140, 140, 146)

        .BackStyle = fmBackStyleTransparent
    End With

    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        "lblAttachmentAlert", _
        True)

    With lbl
        .Left = S(130)
        .Top = S(3)
        .Width = S(280)
        .Height = S(16)

        .Caption = "要確認の添付ファイルあり"
        .Visible = False

        .Font.Name = "Meiryo UI"
        .Font.Size = S(9)
        .Font.Bold = True

        .ForeColor = RGB(176, 64, 60)

        .BackStyle = fmBackStyleTransparent
    End With

    Set lbl = fra.Controls.Add( _
        "Forms.Label.1", _
        "lblAttachmentCount", _
        True)

    With lbl
        .Left = fra.Width - S(58)
        .Top = S(4)
        .Width = S(42)
        .Height = S(16)

        .Caption = LeafFileCount() & "件"

        .Font.Name = "Meiryo UI"
        .Font.Size = S(8)
        .Font.Bold = False

        .ForeColor = RGB(140, 140, 146)

        .TextAlign = fmTextAlignRight

        .BackStyle = fmBackStyleTransparent
    End With

    Set list = fra.Controls.Add( _
        "Forms.Frame.1", _
        "fraAttachments_list", _
        True)

    With list
        .Left = S(1)
        .Top = ATTACH_HEADER
        .Width = fra.Width - S(2) - SCROLL_LANE
        .Height = ATTACH_H - ATTACH_HEADER - S(1)

        .Caption = ""

        .BackColor = RGB(255, 255, 255)

        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
        .ScrollTop = 0
        .ScrollLeft = 0
    End With

    RegisterCardClick fra
    RegisterCardClick fra.Controls("lblAttachmentTitle")
    RegisterCardClick fra.Controls("lblAttachmentAlert")
    RegisterCardClick fra.Controls("lblAttachmentCount")
    RegisterCardClick list

    RenderAttachmentList

End Sub


Public Sub ToggleAttachment(ByVal attachId As String)

    Dim index As Long

    index = FindAttachItem(attachId)
    If index = 0 Then Exit Sub
    If Not mAttachItems(index).Expandable Then Exit Sub

    If AttachItemIsExpanded(attachId) Then
        CollapseAttachment attachId
    Else
        ExpandAttachment attachId
    End If

    RenderAttachmentList

End Sub


Private Sub LoadDummyAttachments()

    mAttachCount = 0
    ExpandAllAttachments

    ' PDF だけ
    AddAttachItem "a1", "", "見積.pdf", "PDF", "pdf", 0, False

    ' 拡張子が変わった ZIP。中は PDF だけ。パスワードは未設定
    AddAttachItem "a2", "", "資料.zi_", "中は PDF  3件", "zipok", 0, True, "none"
    AddAttachItem "a2a", "a2", "表紙.pdf", "PDF", "pdf", 1, False
    AddAttachItem "a2b", "a2", "本編.pdf", "PDF", "pdf", 1, False
    AddAttachItem "a2c", "a2", "別紙.pdf", "PDF", "pdf", 1, False

    ' 入れ子 ZIP。器は判定に出さない。中に PDF 以外がある。パスワードは未設定
    AddAttachItem "a3", "", "提案.zip", "PDF以外あり", "zipbad", 0, True, "none"
    AddAttachItem "a3a", "a3", "cover.pdf", "PDF", "pdf", 1, False
    AddAttachItem "a3b", "a3", "内訳.zip", "器", "container", 1, True, "none"
    AddAttachItem "a3b1", "a3b", "見積.pdf", "PDF", "pdf", 2, False
    AddAttachItem "a3b2", "a3b", "仕様.docx", "PDF以外", "other", 2, False

    ' パスワード付きだが、中身は空
    AddAttachItem "a4", "", "契約.zip", "", "password", 0, True, "set"
    AddAttachItem "a4n", "a4", "中身は空です", "", "note", 1, False

    ' パスワード付き ZIP の中の PDF。これだけ安全
    AddAttachItem "a5", "", "送付.zip", "", "password", 0, True, "set"
    AddAttachItem "a5a", "a5", "本文.pdf", "PDF", "pdf", 1, False

End Sub


Private Sub AddAttachItem( _
    ByVal itemId As String, _
    ByVal parentId As String, _
    ByVal caption As String, _
    ByVal statusText As String, _
    ByVal kind As String, _
    ByVal depth As Long, _
    ByVal expandable As Boolean, _
    Optional ByVal passwordState As String = "")

    mAttachCount = mAttachCount + 1
    ReDim Preserve mAttachItems(1 To mAttachCount)

    With mAttachItems(mAttachCount)
        .Id = itemId
        .ParentId = parentId
        .Caption = caption
        .Status = statusText
        .Kind = kind
        .Depth = depth
        .Expandable = expandable
        .PasswordState = passwordState
    End With

End Sub


Private Sub RenderAttachmentList()

    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim visibleIds() As String
    Dim n As Long
    Dim i As Long
    Dim contentH As Single
    Dim collapsedContentH As Single
    Dim topCount As Long
    Dim attachH As Single
    Dim maxH As Single
    Dim listH As Single
    Dim savedScroll As Single

    Set fra = Me.Controls("fraAttachments")
    Set list = fra.Controls("fraAttachments_list")
    savedScroll = list.ScrollTop

    Set mAttachHandlers = New Collection
    ClearFrameControls list

    n = 0
    ReDim visibleIds(1 To mAttachCount)

    For i = 1 To mAttachCount
        If AttachItemVisible(i) Then
            n = n + 1
            visibleIds(n) = mAttachItems(i).Id
        End If
    Next i

    contentH = n * ATTACH_ROW + S(4)
    topCount = 0
    For i = 1 To mAttachCount
        If Len(mAttachItems(i).ParentId) = 0 Then topCount = topCount + 1
    Next i
    collapsedContentH = topCount * ATTACH_ROW + 4

    If mAttachExpanded Then
        fra.Top = S(30)
        attachH = Me.InsideHeight - MARGIN - S(28) - GAP - fra.Top
    Else
        maxH = Me.InsideHeight - MARGIN - S(28) - GAP - S(30) - GAP - S(150)

        If ATTACH_HEADER + collapsedContentH + S(2) > maxH Then
            attachH = maxH
        Else
            attachH = ATTACH_HEADER + collapsedContentH + S(2)
        End If

        ' 添付欄だけが高くなりすぎないよう、中身に合わせた高さの 2/3 にする
        attachH = attachH * 2 / 3
    End If

    If attachH < ATTACH_HEADER + ATTACH_ROW + S(6) Then
        attachH = ATTACH_HEADER + ATTACH_ROW + S(6)
    End If

    If Not mAttachExpanded Then
        fra.Top = Me.InsideHeight - MARGIN - S(28) - GAP - attachH
    End If

    listH = attachH - ATTACH_HEADER - S(1)

    fra.Height = attachH
    list.Height = listH

    list.ScrollBars = fmScrollBarsNone
    list.ScrollHeight = contentH

    For i = 1 To n
        If i > 1 Then
            If Len(mAttachItems(FindAttachItem(visibleIds(i))).ParentId) = 0 Then
                AddAttachSeparator list, visibleIds(i), (1 + (i - 1) * ATTACH_ROW) - S(1), list.InsideWidth - S(4)
            End If
        End If
        CreateAttachRow list, visibleIds(i), S(1) + (i - 1) * ATTACH_ROW, list.InsideWidth - S(4)
    Next i

    If n > 0 Then
        AddAttachSeparator list, "top", 0, list.InsideWidth - S(4)
        AddAttachSeparator list, "bottom", n * ATTACH_ROW, list.InsideWidth - S(4)
    End If

    If savedScroll < 0 Then savedScroll = 0
    If list.ScrollHeight > list.InsideHeight Then
        If savedScroll > list.ScrollHeight - list.InsideHeight Then
            savedScroll = list.ScrollHeight - list.InsideHeight
        End If
        list.ScrollTop = savedScroll
    Else
        list.ScrollTop = 0
    End If

    If mAttachExpanded Then
        HideRecipientColumns
    Else
        ResizeRecipientColumns fra.Top
    End If

    UpdateAttachmentHeading
    SyncScrollThumb list

End Sub


Public Sub ToggleAttachmentFocus()

    mAttachExpanded = Not mAttachExpanded

    If mAttachExpanded Then
        ExpandAllAttachments
        HideRecipientColumns
    Else
        ApplyCategoryFilter
    End If

    RenderAttachmentList

End Sub


Private Sub HideRecipientColumns()

    Dim names As Variant
    Dim i As Long

    names = Array("fraExternal", "fraOutside", "fraInternal")

    For i = LBound(names) To UBound(names)
        Me.Controls(CStr(names(i))).Visible = False
    Next i

End Sub


Private Sub UpdateAttachmentHeading()

    Dim title As MSForms.Label
    Dim alert As MSForms.Label
    Dim titleWidth As Single

    Set title = Me.Controls("fraAttachments").Controls("lblAttachmentTitle")
    Set alert = Me.Controls("fraAttachments").Controls("lblAttachmentAlert")

    If mAttachExpanded Then
        title.Caption = ChrW(&H25BE) & "  添付ファイル"
        title.ControlTipText = "クリックで宛先カードを戻す"
    Else
        title.Caption = ChrW(&H25B8) & "  添付ファイル"
        title.ControlTipText = "クリックで一覧を上まで広げる"
    End If

    titleWidth = MeasureLabelWidth(title)
    title.AutoSize = False
    title.Width = titleWidth + S(2)
    title.Height = S(18)
    title.Top = S(2)

    alert.Visible = HasAttachmentToReview()
    If Not alert.Visible Then Exit Sub

    alert.AutoSize = False
    alert.Width = MeasureLabelWidth(alert) + S(4)
    alert.AutoSize = False
    alert.Height = S(16)
    alert.Left = title.Left + title.Width + S(10)
    alert.Top = S(3)

End Sub


Private Function HasAttachmentToReview() As Boolean

    Dim i As Long

    For i = 1 To mAttachCount
        If AttachmentNeedsReview(i) Then
            HasAttachmentToReview = True
            Exit Function
        End If
    Next i

End Function


Private Function LeafFileCount() As Long

    Dim i As Long

    For i = 1 To mAttachCount
        If mAttachItems(i).Expandable Then
            ' ZIP と空の ZIP は件数に入れない
        ElseIf mAttachItems(i).Kind <> "note" Then
            LeafFileCount = LeafFileCount + 1
        End If
    Next i

End Function


Private Function AttachmentIsSafe(ByVal index As Long) As Boolean

    Dim parentIndex As Long
    Dim sawPasswordZip As Boolean

    If mAttachItems(index).Kind <> "pdf" Then Exit Function

    parentIndex = FindAttachItem(mAttachItems(index).ParentId)

    Do While parentIndex > 0
        If mAttachItems(parentIndex).PasswordState = "none" Then Exit Function
        If mAttachItems(parentIndex).PasswordState = "set" Then sawPasswordZip = True
        parentIndex = FindAttachItem(mAttachItems(parentIndex).ParentId)
    Loop

    AttachmentIsSafe = sawPasswordZip

End Function


' 要確認は次の行だけ
'   ZIP の外にある PDF
'   パスワード未設定の ZIP
'   PDF 以外のファイル
Private Function AttachmentNeedsReview(ByVal index As Long) As Boolean

    Dim kind As String

    kind = mAttachItems(index).Kind

    If kind = "note" Then Exit Function

    If kind = "other" Then
        AttachmentNeedsReview = True
        Exit Function
    End If

    If kind = "pdf" Then
        If Len(mAttachItems(index).ParentId) = 0 Then
            AttachmentNeedsReview = True
        End If
        Exit Function
    End If

    If mAttachItems(index).Expandable Then
        If mAttachItems(index).PasswordState <> "set" Then
            AttachmentNeedsReview = True
        End If
    End If

End Function


Private Function NearestPasswordZip(ByVal index As Long) As Long

    Dim current As Long

    current = index

    Do While current > 0
        If mAttachItems(current).PasswordState = "set" Then
            NearestPasswordZip = current
            Exit Function
        End If
        current = FindAttachItem(mAttachItems(current).ParentId)
    Loop

End Function


Private Function IsDescendantOf( _
    ByVal index As Long, _
    ByVal ancestorId As String) As Boolean

    Dim parentIndex As Long

    parentIndex = FindAttachItem(mAttachItems(index).ParentId)

    Do While parentIndex > 0
        If mAttachItems(parentIndex).Id = ancestorId Then
            IsDescendantOf = True
            Exit Function
        End If
        parentIndex = FindAttachItem(mAttachItems(parentIndex).ParentId)
    Loop

End Function


' パスワード付き ZIP の中が、空（確認できない注記だけ）か PDF だけか
Private Function ContentsAreEmptyOrPdf(ByVal zipIndex As Long) As Boolean

    Dim i As Long
    Dim kind As String

    ContentsAreEmptyOrPdf = True

    For i = 1 To mAttachCount
        If IsDescendantOf(i, mAttachItems(zipIndex).Id) Then
            kind = mAttachItems(i).Kind
            If kind <> "note" And kind <> "pdf" And Not mAttachItems(i).Expandable Then
                ContentsAreEmptyOrPdf = False
                Exit Function
            End If
        End If
    Next i

End Function


Private Sub AddAttachSeparator( _
    ByVal list As MSForms.Frame, _
    ByVal itemId As String, _
    ByVal topPos As Single, _
    ByVal rowWidth As Single)

    Dim line As MSForms.Label

    Set line = list.Controls.Add( _
        "Forms.Label.1", _
        "ln_" & itemId, _
        True)

    With line
        .Left = 0
        .Top = topPos
        .Width = rowWidth
        .Height = S(1)
        .Caption = ""
        .BackColor = RGB(186, 188, 194)
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
    End With

    line.ZOrder 0

End Sub


Private Sub CreateAttachRow( _
    ByVal list As MSForms.Frame, _
    ByVal itemId As String, _
    ByVal topPos As Single, _
    ByVal rowWidth As Single)

    Dim index As Long
    Dim lbl As MSForms.Label
    Dim nameColor As Long
    Dim statusColor As Long
    Dim backColor As Long
    Dim emphasize As Boolean
    Dim nameLeft As Single
    Dim nameWidth As Single
    Dim statusWidth As Single
    Dim mark As String
    Dim markLeft As Single
    Dim markLbl As MSForms.Label
    Dim nameLbl As MSForms.Label
    Dim statusLbl As MSForms.Label
    Dim pwLbl As MSForms.Label
    Dim cursorLeft As Single
    Dim tailRight As Single

    index = FindAttachItem(itemId)
    If index = 0 Then Exit Sub

    AttachRowStyle index, nameColor, statusColor, backColor
    emphasize = AttachmentNeedsReview(index)

    ' 実行時に作った Frame の中へは部品を置けないので、行はリストへ直接並べる
    Set lbl = list.Controls.Add( _
        "Forms.Label.1", _
        "row_" & itemId, _
        True)

    With lbl
        .Left = 0
        .Top = topPos
        .Width = rowWidth
        .Height = ATTACH_ROW - S(1)
        .Caption = ""
        .BackColor = backColor
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .WordWrap = False
    End With

    lbl.ZOrder 1
    RegisterAttachClick lbl

    If mAttachItems(index).Depth = 0 And mAttachItems(index).Kind <> "note" Then
        Set markLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "ckr_" & itemId, _
            True)

        With markLbl
            .Left = S(4)
            .Top = topPos - S(1)
            .Width = S(16)
            .Height = S(16)
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = S(13)
            .Font.Bold = False
            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
            .Caption = ChrW(&H25CB)
            .ForeColor = RGB(150, 152, 158)
        End With

        markLbl.ZOrder 0
        RegisterAttachClick markLbl

        Set markLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "ck_" & itemId, _
            True)

        With markLbl
            .Left = S(5)
            .Top = topPos + S(3)
            .Width = S(14)
            .Height = S(14)
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = S(11)
            .Font.Bold = False
            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
            .Caption = ChrW(&H25CF)
            .ForeColor = RGB(255, 255, 255)
            .Visible = False
        End With

        markLbl.ZOrder 0
        RegisterAttachClick markLbl

        Set markLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "km_" & itemId, _
            True)

        With markLbl
            .Left = S(4)
            .Top = topPos + S(2)
            .Width = S(16)
            .Height = S(16)
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = S(9)
            .Font.Bold = True
            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
            .Caption = ChrW(&H2713)
            .ForeColor = RGB(32, 33, 36)
            .Visible = False
        End With

        markLbl.ZOrder 0
        RegisterAttachClick markLbl
        PaintAttachCheck list, itemId, mAttachItems(index).Checked
    End If

    ' 左端に余白を取り、ファイル名は赤線を引いたあたりから始める
    markLeft = S(26) + mAttachItems(index).Depth * S(16)
    nameLeft = markLeft + S(14)

    If mAttachItems(index).Expandable Then
        If AttachItemIsExpanded(itemId) Then
            mark = ChrW(&H25BE)
        Else
            mark = ChrW(&H25B8)
        End If

        Set markLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "mk_" & itemId, _
            True)

        With markLbl
            .Left = markLeft
            .Top = topPos + S(2)
            .Width = S(14)
            .Height = S(16)
            .Caption = mark
            .Font.Name = "Meiryo UI"
            .Font.Size = S(9)
            .Font.Bold = emphasize
            .ForeColor = nameColor
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        markLbl.ZOrder 0
        RegisterAttachClick markLbl
    ElseIf mAttachItems(index).Kind <> "note" Then
        Set markLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "fi_" & itemId, _
            True)

        With markLbl
            .Left = markLeft
            .Top = topPos + S(1)
            .Width = S(14)
            .Height = S(16)
            .Caption = ChrW(&HE8A5&)
            .Font.Name = "Segoe MDL2 Assets"
            .Font.Size = S(10)
            .ForeColor = nameColor
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        markLbl.ZOrder 0
        RegisterAttachClick markLbl
    End If

    Set nameLbl = list.Controls.Add( _
        "Forms.Label.1", _
        "att_" & itemId, _
        True)

    With nameLbl
        .Left = nameLeft
        .Top = topPos + S(2)
        .Height = S(16)
        .Caption = mAttachItems(index).Caption
        .Font.Name = "Meiryo UI"
        .Font.Size = S(9)
        .Font.Bold = emphasize
        .Font.Italic = (mAttachItems(index).Kind = "note")
        .ForeColor = nameColor
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .WordWrap = False
        .Width = rowWidth - nameLeft - S(8)
    End With

    nameLbl.ZOrder 0
    RegisterAttachClick nameLbl

    nameWidth = MeasureLabelWidth(nameLbl)
    nameLbl.AutoSize = False
    nameLbl.Height = S(16)
    nameLbl.Top = topPos + S(2)
    nameLbl.Width = nameWidth + S(2)
    cursorLeft = nameLeft + nameLbl.Width + S(12)

    If Len(mAttachItems(index).Status) > 0 Then
        Set statusLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "st_" & itemId, _
            True)

        With statusLbl
            .Top = topPos + S(3)
            .Height = S(16)
            .Caption = mAttachItems(index).Status
            .Font.Name = "Meiryo UI"
            .Font.Size = S(8)
            .Font.Bold = emphasize
            .ForeColor = statusColor
            .TextAlign = fmTextAlignLeft
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        statusWidth = MeasureLabelWidth(statusLbl)
        statusLbl.AutoSize = False
        statusLbl.Height = S(16)
        statusLbl.Top = topPos + S(3)
        statusLbl.Width = statusWidth + S(2)
        statusLbl.Left = cursorLeft
        cursorLeft = statusLbl.Left + statusLbl.Width + S(10)
        statusLbl.ZOrder 0
        RegisterAttachClick statusLbl
    End If

    If Len(PasswordCaption(mAttachItems(index).PasswordState)) > 0 Then
        Set pwLbl = list.Controls.Add( _
            "Forms.Label.1", _
            "pw_" & itemId, _
            True)

        With pwLbl
            .Top = topPos + S(3)
            .Height = S(16)
            .Caption = PasswordCaption(mAttachItems(index).PasswordState)
            .Font.Name = "Meiryo UI"
            .Font.Size = S(8)
            .Font.Bold = emphasize
            .ForeColor = statusColor
            .TextAlign = fmTextAlignLeft
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        statusWidth = MeasureLabelWidth(pwLbl)
        pwLbl.AutoSize = False
        pwLbl.Height = S(16)
        pwLbl.Top = topPos + S(3)
        pwLbl.Width = statusWidth + S(2)
        pwLbl.Left = cursorLeft
        cursorLeft = pwLbl.Left + pwLbl.Width
        pwLbl.ZOrder 0
        RegisterAttachClick pwLbl
    End If

    tailRight = cursorLeft + S(8)
    If tailRight > rowWidth Then
        nameLbl.Width = nameLbl.Width - (tailRight - rowWidth)
        If nameLbl.Width < S(24) Then nameLbl.Width = S(24)
        cursorLeft = nameLeft + nameLbl.Width + S(12)
        If Not statusLbl Is Nothing Then
            statusLbl.Left = cursorLeft
            cursorLeft = statusLbl.Left + statusLbl.Width + S(10)
        End If
        If Not pwLbl Is Nothing Then
            pwLbl.Left = cursorLeft
        End If
    End If

End Sub


Private Function PasswordCaption(ByVal passwordState As String) As String

    Select Case passwordState
        Case "set"
            PasswordCaption = "パスワード設定あり"
        Case "none"
            PasswordCaption = "パスワード設定なし"
    End Select

End Function


Private Sub RegisterAttachClick(ByVal ctl As Object)

    Dim handler As CCardClick

    Set handler = New CCardClick
    Set handler.ParentForm = Me

    If TypeName(ctl) = "Frame" Then
        Set handler.CardFrame = ctl
    Else
        Set handler.CardLabel = ctl
    End If

    mAttachHandlers.Add handler

End Sub


Private Sub ResizeRecipientColumns(ByVal attachTop As Single)

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim areaHeight As Single

    names = Array("fraExternal", "fraOutside", "fraInternal")
    areaHeight = attachTop - GAP - S(30)
    If areaHeight < S(80) Then areaHeight = S(80)

    For i = LBound(names) To UBound(names)
        Set fra = Me.Controls(CStr(names(i)))
        fra.Height = areaHeight
        fra.Controls(fra.Name & "_list").Height = areaHeight - S(34) - S(1)
        SyncScrollThumb fra.Controls(fra.Name & "_list")
    Next i

End Sub


Private Sub ClearFrameControls(ByVal fra As MSForms.Frame)

    Do While fra.Controls.Count > 0
        fra.Controls.Remove 0
    Loop

End Sub


'========================================================
' 細いスクロールバー
'========================================================
Public Sub SyncScrollThumb(ByVal list As Object)

    Dim host As Object
    Dim track As MSForms.Label
    Dim thumb As MSForms.Label
    Dim view As Single
    Dim maxTop As Single
    Dim thumbH As Single
    Dim travel As Single

    If list Is Nothing Then Exit Sub

    Set host = list.Parent
    Set track = FindHostControl(host, "bar_" & list.Name)
    Set thumb = FindHostControl(host, "thumb_" & list.Name)

    If track Is Nothing Or thumb Is Nothing Then
        Set track = host.Controls.Add( _
            "Forms.Label.1", _
            "bar_" & list.Name, _
            True)

        With track
            .Caption = ""
            .BackColor = RGB(214, 216, 220)
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        Set thumb = host.Controls.Add( _
            "Forms.Label.1", _
            "thumb_" & list.Name, _
            True)

        With thumb
            .Caption = ""
            .BackColor = RGB(132, 134, 140)
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleNone
            .WordWrap = False
        End With

        BindScrollDrag CStr(list.Name), track, thumb
    End If

    view = list.InsideHeight
    If view < S(20) Then view = list.Height

    track.Left = list.Left + list.Width + S(2)
    track.Top = list.Top + S(4)
    track.Width = S(6)
    track.BackColor = RGB(214, 216, 220)
    If list.Height > S(16) Then
        track.Height = list.Height - S(12)
    Else
        track.Height = list.Height
    End If

    thumb.Left = track.Left + S(1)
    thumb.Width = S(4)
    thumb.BackColor = RGB(132, 134, 140)

    maxTop = list.ScrollHeight - view

    If maxTop <= S(1) Then
        track.Visible = False
        thumb.Visible = False
        Exit Sub
    End If

    track.Visible = True
    thumb.Visible = True

    thumbH = track.Height * (view / list.ScrollHeight)
    If thumbH < S(28) Then thumbH = S(28)
    If thumbH > track.Height Then thumbH = track.Height
    thumb.Height = thumbH

    travel = track.Height - thumbH
    If list.ScrollTop < 0 Then list.ScrollTop = 0
    If list.ScrollTop > maxTop Then list.ScrollTop = maxTop

    If travel < S(1) Then
        thumb.Top = track.Top
    Else
        thumb.Top = track.Top + (list.ScrollTop / maxTop) * travel
    End If

    track.ZOrder 0
    thumb.ZOrder 0

End Sub


Public Sub MoveScrollThumb( _
    ByVal listName As String, _
    ByVal delta As Single)

    Dim list As Object
    Dim track As MSForms.Label
    Dim thumb As MSForms.Label
    Dim view As Single
    Dim maxTop As Single
    Dim travel As Single
    Dim newTop As Single

    Set list = ScrollListByName(listName)
    If list Is Nothing Then Exit Sub

    Set track = list.Parent.Controls("bar_" & listName)
    Set thumb = list.Parent.Controls("thumb_" & listName)

    view = list.InsideHeight
    If view < S(20) Then view = list.Height
    maxTop = list.ScrollHeight - view
    travel = track.Height - thumb.Height
    If maxTop <= S(1) Or travel < S(1) Then Exit Sub

    newTop = thumb.Top + delta
    If newTop < track.Top Then newTop = track.Top
    If newTop > track.Top + travel Then newTop = track.Top + travel

    thumb.Top = newTop
    list.ScrollTop = (newTop - track.Top) / travel * maxTop

End Sub


Public Sub JumpScrollThumb( _
    ByVal listName As String, _
    ByVal y As Single)

    Dim list As Object
    Dim track As MSForms.Label
    Dim thumb As MSForms.Label
    Dim delta As Single

    Set list = ScrollListByName(listName)
    If list Is Nothing Then Exit Sub

    Set track = list.Parent.Controls("bar_" & listName)
    Set thumb = list.Parent.Controls("thumb_" & listName)

    delta = (track.Top + y - thumb.Height / 2) - thumb.Top
    MoveScrollThumb listName, delta

End Sub


Private Sub BindScrollDrag( _
    ByVal listName As String, _
    ByVal track As MSForms.Label, _
    ByVal thumb As MSForms.Label)

    Dim handler As CScrollDrag
    Dim existing As Object

    On Error Resume Next
    Set existing = mScrollDrags.Item(listName)
    On Error GoTo 0
    If Not existing Is Nothing Then Exit Sub

    Set handler = New CScrollDrag
    Set handler.Track = track
    Set handler.Thumb = thumb
    Set handler.ParentForm = Me
    handler.ListName = listName
    mScrollDrags.Add handler, listName

End Sub


Private Function ScrollListByName(ByVal listName As String) As Object

    Dim hostName As String

    If Len(listName) <= 5 Then Exit Function
    hostName = Left$(listName, Len(listName) - 5)

    On Error Resume Next
    Set ScrollListByName = Me.Controls(hostName).Controls(listName)

End Function


Private Function FindHostControl( _
    ByVal host As Object, _
    ByVal controlName As String) As Object

    On Error Resume Next
    Set FindHostControl = host.Controls(controlName)

End Function


Private Function AttachItemVisible(ByVal index As Long) As Boolean

    Dim parentIndex As Long

    If Len(mAttachItems(index).ParentId) = 0 Then
        AttachItemVisible = True
        Exit Function
    End If

    parentIndex = FindAttachItem(mAttachItems(index).ParentId)
    If parentIndex = 0 Then Exit Function

    AttachItemVisible = AttachItemVisible(parentIndex) And AttachItemIsExpanded(mAttachItems(index).ParentId)

End Function


Private Function AttachItemIsExpanded(ByVal itemId As String) As Boolean

    Dim saved As String

    On Error Resume Next
    saved = CStr(mCollapsedIds(itemId))
    If Err.Number <> 0 Then
        AttachItemIsExpanded = True
        Err.Clear
    End If

End Function


Private Sub ExpandAllAttachments()

    Set mCollapsedIds = New Collection

End Sub


Private Sub CollapseAttachment(ByVal attachId As String)

    On Error Resume Next
    mCollapsedIds.Remove attachId
    On Error GoTo 0
    mCollapsedIds.Add attachId, attachId

End Sub


Private Sub ExpandAttachment(ByVal attachId As String)

    On Error Resume Next
    mCollapsedIds.Remove attachId
    On Error GoTo 0

End Sub


Private Function FindAttachItem(ByVal itemId As String) As Long

    Dim i As Long

    For i = 1 To mAttachCount
        If mAttachItems(i).Id = itemId Then
            FindAttachItem = i
            Exit Function
        End If
    Next i

End Function


Private Sub AttachRowStyle( _
    ByVal index As Long, _
    ByRef nameColor As Long, _
    ByRef statusColor As Long, _
    ByRef backColor As Long)

    Dim zipIndex As Long

    nameColor = RGB(55, 55, 60)
    statusColor = RGB(55, 55, 60)
    backColor = RGB(255, 255, 255)

    ' 安全なのは、パスワード付き ZIP の中にある PDF だけ
    If AttachmentIsSafe(index) Then
        nameColor = RGB(70, 120, 78)
        statusColor = RGB(70, 120, 78)
        Exit Sub
    End If

    ' 見出しの「要確認」と同じ赤
    If AttachmentNeedsReview(index) Then
        nameColor = RGB(176, 64, 60)
        statusColor = RGB(176, 64, 60)
    End If

    ' パスワード付きで、中身が空または PDF だけなら赤背景にしない
    zipIndex = NearestPasswordZip(index)
    If zipIndex > 0 Then
        If ContentsAreEmptyOrPdf(zipIndex) Then Exit Sub
    End If

    backColor = RGB(252, 244, 243)

End Sub


Private Sub PlaceInternalCheckButton()

    Dim fra As MSForms.Frame
    Dim titleLbl As MSForms.Label
    Dim textW As Single

    Set fra = Me.Controls("fraInternal")
    Set titleLbl = fra.Controls("fraInternal_title")

    titleLbl.AutoSize = True
    textW = titleLbl.Width
    titleLbl.AutoSize = False
    titleLbl.Width = textW
    titleLbl.Height = S(22)
    titleLbl.Top = S(8)

    CreateFlatButton _
        "btnCheckInternal", _
        "まとめてチェック", _
        "", _
        titleLbl.Left + textW + S(6), _
        S(8), _
        S(100), _
        S(18), _
        RGB(255, 255, 255), _
        RGB(42, 96, 148), _
        RGB(186, 204, 222), _
        RGB(232, 238, 244), _
        False, _
        "CheckInternal", _
        fra, _
        S(8)
    SetControlTip fra.Controls("btnCheckInternal"), _
        "部内の人だけ、まとめてチェックを付けます"

End Sub


'========================================================
' フッター
'========================================================
Private Sub CreateFooter()

    Dim topPos As Single
    Dim sendLeft As Single
    Dim x As Single
    Dim scaleCaption As String

    topPos = Me.InsideHeight - MARGIN - S(28)
    sendLeft = Me.InsideWidth - MARGIN - S(96)
    x = MARGIN

    '表示の初期化。枠だけの静かなボタンにして、送信と並べない
    CreateFlatButton _
        "btnReset", _
        "表示を初期化", _
        ChrW(&H21BA), _
        x, _
        topPos, _
        S(112), _
        S(28), _
        RGB(255, 255, 255), _
        RGB(90, 92, 98), _
        RGB(214, 216, 220), _
        RGB(236, 237, 240), _
        False, _
        "Reset"
    SetControlTip Me.Controls("btnReset"), _
        "列の絞り込み、開いているカード、ZIP の展開、スクロールを最初の状態に戻します"
    x = x + S(112) + S(6)

    CreateFlatButton _
        "btnCheckAll", _
        "チェックを全部付ける（非推奨）", _
        "", _
        x, _
        topPos, _
        S(128), _
        S(28), _
        RGB(255, 255, 255), _
        RGB(140, 142, 148), _
        RGB(214, 216, 220), _
        RGB(236, 237, 240), _
        False, _
        "CheckAll", _
        fontSize:=S(8)
    SetControlTip Me.Controls("btnCheckAll"), _
        "全員と、添付ファイルの1階層目にチェックを付けます"
    x = x + S(128) + S(6)

    CreateFlatButton _
        "btnClearChecks", _
        "チェックを全部外す", _
        "", _
        x, _
        topPos, _
        S(124), _
        S(28), _
        RGB(255, 255, 255), _
        RGB(90, 92, 98), _
        RGB(214, 216, 220), _
        RGB(236, 237, 240), _
        False, _
        "ClearChecks"
    SetControlTip Me.Controls("btnClearChecks"), _
        "人のチェックと、添付ファイルの1階層目のチェックをすべて外します"

    x = x + S(124) + S(6)

    If mScale > 1 Then
        scaleCaption = ChrW(&H6A19) & ChrW(&H6E96)
    Else
        scaleCaption = "1.5" & ChrW(&H500D)
    End If

    CreateFlatButton _
        "btnScale", _
        scaleCaption, _
        "", _
        x, _
        topPos, _
        S(64), _
        S(28), _
        RGB(255, 255, 255), _
        RGB(45, 45, 48), _
        RGB(214, 216, 220), _
        RGB(236, 237, 240), _
        False, _
        "Scale"
    SetControlTip Me.Controls("btnScale"), ScaleTip()
    CreateFlatButton _
        "btnCancel", _
        "キャンセル", _
        "", _
        sendLeft - S(8) - S(96), _
        topPos, _
        S(96), _
        S(28), _
        RGB(255, 255, 255), _
        RGB(45, 45, 48), _
        RGB(214, 216, 220), _
        RGB(236, 237, 240), _
        False, _
        "Cancel"

    CreateFlatButton _
        "btnSend", _
        "送信", _
        "", _
        sendLeft, _
        topPos, _
        S(96), _
        S(28), _
        RGB(32, 33, 36), _
        RGB(255, 255, 255), _
        RGB(32, 33, 36), _
        RGB(16, 16, 18), _
        True, _
        "Send"

    UpdateSendButton

End Sub


Private Sub CheckPeopleIn(ByVal frameName As String)

    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim ctl As Object
    Dim key As String

    Set fra = Me.Controls(frameName)
    Set list = fra.Controls(fra.Name & "_list")

    For Each ctl In list.Controls
        If TypeName(ctl) = "Frame" Then
            If Left$(ctl.Name, 5) = "card_" Then
                If Not PersonIsChecked(ctl) Then
                    key = PersonKey(ctl)
                    mCheckedPeople.Add True, key
                End If
                PlacePersonCheck ctl
            End If
        End If
    Next ctl

End Sub


Private Sub CheckAllAttachments()

    Dim i As Long
    Dim list As MSForms.Frame
    Dim ctl As Object

    For i = 1 To mAttachCount
        If mAttachItems(i).Depth = 0 Then
            If mAttachItems(i).Kind <> "note" Then
                mAttachItems(i).Checked = True
            End If
        End If
    Next i

    Set list = Me.Controls("fraAttachments").Controls("fraAttachments_list")
    For i = 1 To mAttachCount
        If mAttachItems(i).Depth = 0 Then
            If mAttachItems(i).Kind <> "note" Then
                PaintAttachCheck list, mAttachItems(i).Id, True
            End If
        End If
    Next i

End Sub


Private Sub CheckAll()

    CheckPeopleIn "fraExternal"
    CheckPeopleIn "fraOutside"
    CheckPeopleIn "fraInternal"
    CheckAllAttachments
    UpdateSendButton

End Sub


Private Sub CheckInternal()

    CheckPeopleIn "fraInternal"
    UpdateSendButton

End Sub


'========================================================
' フラットなフッターボタン
'========================================================
Private Sub CreateFlatButton( _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal iconChar As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal width As Single, _
    ByVal height As Single, _
    ByVal backColor As Long, _
    ByVal foreColor As Long, _
    ByVal borderColor As Long, _
    ByVal downColor As Long, _
    ByVal bold As Boolean, _
    ByVal actionName As String, _
    Optional ByVal host As MSForms.Frame = Nothing, _
    Optional ByVal fontSize As Single = 0)

    Dim plate As MSForms.Frame
    Dim face As MSForms.Label
    Dim icon As MSForms.Label

    If fontSize <= 0 Then fontSize = S(9)

    If host Is Nothing Then
        Set plate = Me.Controls.Add( _
            "Forms.Frame.1", _
            controlName, _
            True)
    Else
        Set plate = host.Controls.Add( _
            "Forms.Frame.1", _
            controlName, _
            True)
    End If

    With plate
        .Left = leftPos
        .Top = topPos
        .Width = width
        .Height = height

        .Caption = ""

        .BackColor = backColor
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = borderColor
        .SpecialEffect = fmSpecialEffectFlat

        .ScrollBars = fmScrollBarsNone
    End With

    If Len(iconChar) > 0 Then
        Set icon = plate.Controls.Add( _
            "Forms.Label.1", _
            controlName & "_icon", _
            True)

        With icon
            .Caption = iconChar
            .AutoSize = False
            .WordWrap = False
            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
            .ForeColor = foreColor
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = S(11)
        End With
    End If

    Set face = plate.Controls.Add( _
        "Forms.Label.1", _
        controlName & "_face", _
        True)

    With face
        .Caption = caption
        .AutoSize = False
        .WordWrap = False
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
        .ForeColor = foreColor
        .Font.Name = "Meiryo UI"
        .Font.Size = fontSize
        .Font.Bold = bold
    End With

    CenterButtonCaption plate, face, icon

    If Not icon Is Nothing Then
        RegisterFlatClick icon, plate, backColor, downColor, actionName
    End If
    RegisterFlatClick plate, plate, backColor, downColor, actionName
    RegisterFlatClick face, plate, backColor, downColor, actionName

End Sub


' 文字の実寸を測って、ボタンの内側の中央に置く
Private Sub CenterButtonCaption( _
    ByVal plate As MSForms.Frame, _
    ByVal face As MSForms.Label, _
    ByVal icon As MSForms.Label)

    Dim gap As Single
    Dim groupW As Single
    Dim groupLeft As Single
    Dim textW As Single
    Dim iconW As Single
    Dim boxW As Single
    Dim boxH As Single

    gap = S(4)
    boxW = plate.Width - S(2)
    boxH = plate.Height - S(2)

    textW = MeasureLabelWidth(face) + S(8)
    face.TextAlign = fmTextAlignCenter
    face.Width = textW
    face.Height = face.Font.Size + S(6)
    face.Top = (boxH - face.Font.Size) / 2

    If icon Is Nothing Then
        face.Left = (boxW - textW) / 2
        Exit Sub
    End If

    iconW = MeasureLabelWidth(icon) + S(2)
    icon.TextAlign = fmTextAlignCenter
    icon.Width = iconW
    icon.Height = icon.Font.Size + S(6)
    icon.Top = (boxH - icon.Font.Size) / 2

    groupW = iconW + gap + textW
    groupLeft = (boxW - groupW) / 2
    If groupLeft < S(2) Then groupLeft = S(2)

    icon.Left = groupLeft
    face.Left = groupLeft + iconW + gap

End Sub


Private Function MeasureLabelWidth(ByVal lbl As MSForms.Label) As Single

    On Error GoTo Fallback

    lbl.WordWrap = False
    lbl.AutoSize = False
    lbl.Width = S(200)
    lbl.Height = S(24)
    lbl.AutoSize = True
    MeasureLabelWidth = lbl.Width
    lbl.AutoSize = False

    If MeasureLabelWidth < S(8) Then MeasureLabelWidth = S(8)
    Exit Function

Fallback:
    MeasureLabelWidth = lbl.Font.Size * Len(lbl.Caption) + S(4)
    lbl.AutoSize = False

End Function


'========================================================
' Buttonイベント登録
'========================================================
Private Sub RegisterFlatClick( _
    ByVal ctl As Object, _
    ByVal host As Object, _
    ByVal idleBack As Long, _
    ByVal downBack As Long, _
    ByVal actionName As String)

    Dim handler As CButtonHandler

    Set handler = New CButtonHandler

    Set handler.ParentForm = Me
    Set handler.Host = host
    handler.ActionName = actionName
    handler.IdleBack = idleBack
    handler.DownBack = downBack

    If TypeName(ctl) = "Frame" Then
        Set handler.Plate = ctl
    Else
        Set handler.Face = ctl
    End If

    mButtonHandlers.Add handler

End Sub


'========================================================
' プロフィール画像の確認チェック
'========================================================
Private Function PersonKey(ByVal card As MSForms.Frame) As String

    PersonKey = card.Parent.Name & "|" & card.Name

End Function


Private Function PersonIsChecked(ByVal card As MSForms.Frame) As Boolean

    On Error Resume Next
    PersonIsChecked = False
    If mCheckedPeople Is Nothing Then Exit Function
    PersonIsChecked = CBool(mCheckedPeople(PersonKey(card)))
    If Err.Number <> 0 Then
        Err.Clear
        PersonIsChecked = False
    End If

End Function


Private Sub PlacePersonCheck(ByVal card As MSForms.Frame)

    Dim cardNo As String
    Dim photo As MSForms.Label
    Dim chk As MSForms.Label
    Dim plate As MSForms.Label
    Dim mark As Single

    cardNo = Mid$(card.Name, 6)
    Set photo = card.Controls("photo_" & cardNo)
    Set plate = card.Controls("phshd_" & cardNo)
    Set chk = card.Controls("phchk_" & cardNo)

    photo.Font.Name = "Meiryo UI"
    photo.Caption = CStr(photo.Tag)

    mark = photo.Left - S(3)
    If mark > S(16) Then mark = S(16)
    If mark > photo.Height - S(2) Then mark = photo.Height - S(2)
    If mark < S(11) Then mark = S(11)

    With plate
        .Width = mark
        .Height = mark + S(4)
        .Left = photo.Left - .Width - S(1)
        If .Left < 0 Then .Left = 0
        .Top = photo.Top + (photo.Height - mark) / 2
        If .Top < S(14) Then .Top = S(14)
        .Caption = ChrW(&H25CF)
        .Font.Name = "Segoe UI Symbol"
        .Font.Bold = False
        .Font.Size = mark * 0.92
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .Visible = True
    End With

    With chk
        .Width = mark * 0.78
        .Height = mark * 0.72
        .Left = plate.Left + (mark - .Width) / 2 - mark * 0.02
        .Top = plate.Top + (mark - .Height) / 2 + mark * 0.08
        .Font.Name = "Segoe UI Symbol"
        .Font.Bold = True
        .Font.Size = mark * 0.56
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .Caption = ChrW(&H2713)
        .TextAlign = fmTextAlignCenter
        .Visible = True
    End With

    If PersonIsChecked(card) Then
        plate.ForeColor = RGB(46, 164, 90)
        chk.ForeColor = RGB(255, 255, 255)
    Else
        plate.ForeColor = RGB(198, 200, 204)
        chk.ForeColor = RGB(132, 134, 138)
    End If
    plate.ZOrder 0
    chk.ZOrder 0

End Sub


Public Sub TogglePersonCheck(ByVal card As MSForms.Frame)

    Dim key As String

    If card Is Nothing Then Exit Sub
    If Left$(card.Name, 5) <> "card_" Then Exit Sub

    key = PersonKey(card)
    If PersonIsChecked(card) Then
        On Error Resume Next
        mCheckedPeople.Remove key
        On Error GoTo 0
    Else
        mCheckedPeople.Add True, key
    End If

    PlacePersonCheck card
    UpdateSendButton

End Sub


Private Sub PaintAttachCheck( _
    ByVal list As MSForms.Frame, _
    ByVal itemId As String, _
    ByVal checked As Boolean)

    Dim ring As MSForms.Label
    Dim disc As MSForms.Label
    Dim tick As MSForms.Label

    On Error Resume Next
    Set ring = list.Controls("ckr_" & itemId)
    Set disc = list.Controls("ck_" & itemId)
    Set tick = list.Controls("km_" & itemId)
    On Error GoTo 0
    If ring Is Nothing Or disc Is Nothing Or tick Is Nothing Then Exit Sub

    ring.Caption = ChrW(&H25CF)
    ring.Font.Name = "Segoe UI Symbol"
    ring.Font.Bold = False
    ring.Font.Size = ring.Width * 0.92
    ring.Visible = True

    disc.Visible = False

    tick.Caption = ChrW(&H2713)
    tick.Font.Name = "Segoe UI Symbol"
    tick.Font.Bold = True
    tick.Font.Size = ring.Width * 0.56
    tick.Width = ring.Width * 0.78
    tick.Height = ring.Width * 0.72
    tick.Left = ring.Left + (ring.Width - tick.Width) / 2 - ring.Width * 0.02
    tick.Top = ring.Top + (ring.Width - tick.Height) / 2 + ring.Width * 0.08
    tick.Visible = True

    If checked Then
        ring.ForeColor = RGB(46, 164, 90)
        tick.ForeColor = RGB(255, 255, 255)
    Else
        ring.ForeColor = RGB(198, 200, 204)
        tick.ForeColor = RGB(132, 134, 138)
    End If
    ring.ZOrder 0
    tick.ZOrder 0

End Sub

Public Sub ToggleAttachCheck(ByVal itemId As String)

    Dim index As Long
    Dim lbl As MSForms.Label

    index = FindAttachItem(itemId)
    If index = 0 Then Exit Sub
    If mAttachItems(index).Depth <> 0 Then Exit Sub
    If mAttachItems(index).Kind = "note" Then Exit Sub

    mAttachItems(index).Checked = Not mAttachItems(index).Checked

    PaintAttachCheck _
        Me.Controls("fraAttachments").Controls("fraAttachments_list"), _
        itemId, _
        mAttachItems(index).Checked

    UpdateSendButton

End Sub


Private Function AllPeopleChecked() As Boolean

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim ctl As Object
    Dim found As Boolean

    names = Array("fraExternal", "fraOutside", "fraInternal")
    found = False

    For i = LBound(names) To UBound(names)
        Set fra = Me.Controls(CStr(names(i)))
        Set list = fra.Controls(fra.Name & "_list")
        For Each ctl In list.Controls
            If TypeName(ctl) = "Frame" Then
                If Left$(ctl.Name, 5) = "card_" Then
                    found = True
                    If Not PersonIsChecked(ctl) Then Exit Function
                End If
            End If
        Next ctl
    Next i

    AllPeopleChecked = found

End Function


Private Function AllTopFilesChecked() As Boolean

    Dim i As Long
    Dim found As Boolean

    found = False
    For i = 1 To mAttachCount
        If mAttachItems(i).Depth = 0 Then
            If mAttachItems(i).Kind <> "note" Then
                found = True
                If Not mAttachItems(i).Checked Then Exit Function
            End If
        End If
    Next i

    AllTopFilesChecked = found

End Function


Private Sub SetControlTip(ByVal ctl As Object, ByVal tip As String)

    Dim child As Object

    On Error Resume Next
    ctl.ControlTipText = tip
    If TypeName(ctl) = "Frame" Then
        For Each child In ctl.Controls
            child.ControlTipText = tip
        Next child
    End If
    On Error GoTo 0

End Sub


Private Sub ClearAllChecks()

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim ctl As Object

    Set mCheckedPeople = New Collection

    names = Array("fraExternal", "fraOutside", "fraInternal")
    For i = LBound(names) To UBound(names)
        Set fra = Me.Controls(CStr(names(i)))
        Set list = fra.Controls(fra.Name & "_list")
        For Each ctl In list.Controls
            If TypeName(ctl) = "Frame" Then
                If Left$(ctl.Name, 5) = "card_" Then
                    PlacePersonCheck ctl
                End If
            End If
        Next ctl
    Next i

    For i = 1 To mAttachCount
        mAttachItems(i).Checked = False
    Next i

    Set list = Me.Controls("fraAttachments").Controls("fraAttachments_list")
    For i = 1 To mAttachCount
        If mAttachItems(i).Depth = 0 Then
            If mAttachItems(i).Kind <> "note" Then
                PaintAttachCheck list, mAttachItems(i).Id, False
            End If
        End If
    Next i

    UpdateSendButton

End Sub


Private Function RemainCaption(ByVal count As Long) As String

    Dim head As String
    Dim tail As String

    tail = " " & ChrW(&H672A) & ChrW(&H2713) & ChrW(&H4EF6) & ChrW(&H6570) & ChrW(&HFF1A) & CStr(count)

    If count = 0 Then
        head = ChrW(&H9001) & ChrW(&H4FE1) & ChrW(&H3067) & ChrW(&H304D) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002)
    Else
        head = ChrW(&H5B9B) & ChrW(&H5148) & ChrW(&H3001) & ChrW(&H6DFB) & ChrW(&H4ED8) & _
            ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & _
            ChrW(&H3092) & ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF) & _
            ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
    End If

    RemainCaption = head & tail

End Function


Private Function UncheckedRequiredCount() As Long

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim ctl As Object
    Dim n As Long

    names = Array("fraExternal", "fraOutside", "fraInternal")

    For i = LBound(names) To UBound(names)
        Set fra = Me.Controls(CStr(names(i)))
        Set list = fra.Controls(fra.Name & "_list")
        For Each ctl In list.Controls
            If TypeName(ctl) = "Frame" Then
                If Left$(ctl.Name, 5) = "card_" Then
                    If Not PersonIsChecked(ctl) Then n = n + 1
                End If
            End If
        Next ctl
    Next i

    For i = 1 To mAttachCount
        If mAttachItems(i).Depth = 0 Then
            If mAttachItems(i).Kind <> "note" Then
                If Not mAttachItems(i).Checked Then n = n + 1
            End If
        End If
    Next i

    UncheckedRequiredCount = n

End Function


Private Sub UpdateRemainLabel()

    Dim lbl As MSForms.Label
    Dim blank As MSForms.Label
    Dim blankRight As MSForms.Label
    Dim n As Long

    On Error Resume Next
    Set lbl = Me.Controls("lblCheckRemain")
    If lbl Is Nothing Then Exit Sub
    On Error GoTo 0

    n = UncheckedRequiredCount()
    lbl.Caption = RemainCaption(n)
    If n = 0 Then
        lbl.ForeColor = RGB(46, 164, 90)
    Else
        lbl.ForeColor = RGB(176, 64, 60)
    End If

    lbl.WordWrap = False
    lbl.AutoSize = False
    lbl.TextAlign = fmTextAlignCenter
    lbl.Width = MeasureLabelWidth(lbl) + S(2)
    lbl.Height = S(20)

    On Error Resume Next
    Set blank = Me.Controls("lblHeaderBlank")
    Set blankRight = Me.Controls("lblHeaderBlankRight")
    On Error GoTo 0

    lbl.Left = (Me.InsideWidth - lbl.Width) / 2
    If lbl.Left < 0 Then lbl.Left = 0
    If lbl.Left + lbl.Width > Me.InsideWidth Then
        lbl.Left = Me.InsideWidth - lbl.Width
    End If

    If Not blank Is Nothing Then
        If lbl.Left > blank.Left Then
            blank.Width = lbl.Left - blank.Left
        Else
            blank.Width = S(1)
        End If
    End If

    If Not blankRight Is Nothing Then
        blankRight.Top = 0
        blankRight.Height = S(28)
        blankRight.Left = lbl.Left + lbl.Width
        If blankRight.Left < Me.InsideWidth Then
            blankRight.Width = Me.InsideWidth - blankRight.Left
        Else
            blankRight.Left = Me.InsideWidth - S(1)
            blankRight.Width = S(1)
        End If
    End If

    lbl.ZOrder 0

End Sub

Private Sub UpdateSendButton()

    Dim ready As Boolean
    Dim plate As MSForms.Frame
    Dim face As MSForms.Label
    Dim backColor As Long
    Dim foreColor As Long
    Dim downColor As Long
    Dim handler As Variant

    On Error Resume Next
    Set plate = Me.Controls("btnSend")
    If plate Is Nothing Then Exit Sub
    Set face = plate.Controls("btnSend_face")
    On Error GoTo 0

    ready = AllPeopleChecked() And AllTopFilesChecked()

    If ready Then
        backColor = RGB(32, 33, 36)
        foreColor = RGB(255, 255, 255)
        downColor = RGB(16, 16, 18)
        plate.Tag = "on"
        SetControlTip plate, ""
    Else
        backColor = RGB(196, 198, 202)
        foreColor = RGB(255, 255, 255)
        downColor = backColor
        plate.Tag = "off"
        SetControlTip plate, _
            "人の画像と添付の1階層目に、チェックを全部付けないと送信できません"
    End If

    plate.BackColor = backColor
    plate.BorderColor = backColor
    If Not face Is Nothing Then face.ForeColor = foreColor

    For Each handler In mButtonHandlers
        If handler.ActionName = "Send" Then
            handler.IdleBack = backColor
            handler.DownBack = downColor
        End If
    Next handler

    UpdateRemainLabel

End Sub


'========================================================
' ボタンから呼ばれる処理
'========================================================
Public Sub ButtonAction(ByVal actionName As String)

    Select Case actionName

        Case "Cancel"

            Unload Me

        Case "Send"

            If Me.Controls("btnSend").Tag <> "on" Then Exit Sub

            MsgBox _
                "これは試作版です。" & vbCrLf & vbCrLf & _
                "次の段階で、実際のOutlookメールを送信する処理を追加します。", _
                vbInformation, _
                "送信前チェック"

        Case "Reset"

            ResetDisplay

        Case "ClearChecks"

            ClearAllChecks

        Case "CheckAll"

            CheckAll

        Case "CheckInternal"

            CheckInternal

        Case "Scale"

            ApplyDisplayScale

    End Select

End Sub


'========================================================
' 画面の開き方だけを、起動直後の状態に戻す
'========================================================
Private Sub ResetDisplay()

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame

    If Not mExpandedCard Is Nothing Then
        ApplyCardLayout mExpandedCard, False
        Set mExpandedCard = Nothing
    End If

    mCategoryFilter = ""
    mAddressFilter = ""
    mAttachExpanded = False
    ApplyCategoryFilter
    ApplyAddressFilter

    names = Array("fraExternal", "fraOutside", "fraInternal")

    For i = LBound(names) To UBound(names)
        Set fra = Me.Controls(CStr(names(i)))
        Set list = fra.Controls(fra.Name & "_list")
        list.ScrollTop = 0
        SyncScrollThumb list
    Next i

    ExpandAllAttachments
    RenderAttachmentList

    Set list = Me.Controls("fraAttachments").Controls("fraAttachments_list")
    list.ScrollTop = 0

End Sub


'========================================================
' イニシャル取得
'========================================================
Private Function GetInitials(ByVal fullName As String) As String

    Dim parts() As String

    fullName = Trim(fullName)

    If InStr(fullName, " ") > 0 Then

        parts = Split(fullName, " ")

        If UBound(parts) >= 1 Then
            GetInitials = Left(parts(0), 1) & Left(parts(1), 1)
        Else
            GetInitials = Left(fullName, 2)
        End If

    Else

        GetInitials = Left(fullName, 2)

    End If

End Function


Private Function ScaleTip() As String

    ScaleTip = ChrW(&H753B) & ChrW(&H9762) & ChrW(&H5168) & ChrW(&H4F53) & ChrW(&H3092) & _
        "1.5" & ChrW(&H500D) & ChrW(&H306B) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & _
        ChrW(&H3002) & ChrW(&H53CE) & ChrW(&H307E) & ChrW(&H3089) & ChrW(&H306A) & ChrW(&H3044) & _
        ChrW(&H74B0) & ChrW(&H5883) & ChrW(&H3067) & ChrW(&H306F) & ChrW(&H5207) & ChrW(&H308A) & _
        ChrW(&H66FF) & ChrW(&H3048) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093)

End Function


Private Sub RememberChecks()

    Dim files As Collection
    Dim i As Long

    Set files = New Collection
    For i = 1 To mAttachCount
        If mAttachItems(i).Checked Then files.Add mAttachItems(i).Id
    Next i
    KeepSendCheckState mCheckedPeople, files

End Sub


Private Sub RestoreChecks()

    Dim names As Variant
    Dim i As Long
    Dim fra As MSForms.Frame
    Dim list As MSForms.Frame
    Dim ctl As Object
    Dim key As Variant
    Dim id As Variant
    Dim index As Long

    If KeptPeople Is Nothing And KeptFiles Is Nothing Then Exit Sub

    If Not KeptPeople Is Nothing Then
        Set mCheckedPeople = New Collection
        For Each key In KeptPeople
            On Error Resume Next
            mCheckedPeople.Add True, CStr(key)
            On Error GoTo 0
        Next key
        names = Array("fraExternal", "fraOutside", "fraInternal")
        For i = LBound(names) To UBound(names)
            Set fra = Me.Controls(CStr(names(i)))
            Set list = fra.Controls(fra.Name & "_list")
            For Each ctl In list.Controls
                If TypeName(ctl) = "Frame" Then
                    If Left$(ctl.Name, 5) = "card_" Then PlacePersonCheck ctl
                End If
            Next ctl
        Next i
    End If

    If Not KeptFiles Is Nothing Then
        For Each id In KeptFiles
            index = FindAttachItem(CStr(id))
            If index > 0 Then
                mAttachItems(index).Checked = True
                PaintAttachCheck _
                    Me.Controls("fraAttachments").Controls("fraAttachments_list"), _
                    CStr(id), _
                    True
            End If
        Next id
    End If

    Set KeptPeople = Nothing
    Set KeptFiles = Nothing
    UpdateSendButton

End Sub


Private Sub ApplyDisplayScale()

    Dim nextScale As Single

    If mScale > 1 Then
        nextScale = 1
    Else
        nextScale = 1.5
    End If

    If Not SendCheckUiFits(nextScale) Then
        MsgBox _
            ChrW(&H3053) & ChrW(&H306E) & ChrW(&H753B) & ChrW(&H9762) & ChrW(&H3067) & _
            ChrW(&H306F) & " 1.5 " & ChrW(&H500D) & ChrW(&H3067) & ChrW(&H306F) & _
            ChrW(&H53CE) & ChrW(&H307E) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & _
            ChrW(&H3093) & ChrW(&H3002), _
            vbExclamation, _
            ChrW(&H9001) & ChrW(&H4FE1) & ChrW(&H524D) & ChrW(&H30C1) & ChrW(&H30A7) & _
            ChrW(&H30C3) & ChrW(&H30AF)
        Exit Sub
    End If

    RememberChecks
    SendCheckScale = nextScale
    Application.OnTime Now + TimeSerial(0, 0, 1), "ReopenSendCheckScaled"
    Unload Me

End Sub

'========================================================
' Label作成ヘルパー
'========================================================

Private Function AddLabel( _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal width As Single, _
    ByVal height As Single) As MSForms.Label

    Dim lbl As MSForms.Label

    Set lbl = Me.Controls.Add( _
        "Forms.Label.1", _
        controlName, _
        True)

    With lbl
        .Left = leftPos
        .Top = topPos
        .Width = width
        .Height = height
        .Caption = caption
    End With

    Set AddLabel = lbl

End Function
