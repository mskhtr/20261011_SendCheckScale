VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCardClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents CardFrame As MSForms.Frame
Public WithEvents CardLabel As MSForms.Label

Public ParentForm As Object


Private Sub CardFrame_Click()

    If ParentForm Is Nothing Then Exit Sub
    If CardFrame Is Nothing Then Exit Sub

    If CardFrame.Name = "fraAttachments" Or CardFrame.Name = "fraAttachments_list" Then
        ParentForm.ToggleAttachmentFocus
    ElseIf Left(CardFrame.Name, 4) = "att_" Then
        ParentForm.ToggleAttachment Mid$(CardFrame.Name, 5)
    ElseIf Left(CardFrame.Name, 5) = "card_" Then
        ParentForm.ToggleCard CardFrame
    Else
        ParentForm.ToggleCategoryFilter CardFrame
    End If

End Sub


Private Sub CardLabel_Click()

    Dim card As Object

    If ParentForm Is Nothing Then Exit Sub
    If CardLabel Is Nothing Then Exit Sub

    If CardLabel.Name = "roleTo" Or _
        CardLabel.Name = "roleCc" Or _
        CardLabel.Name = "roleBcc" Then
        ParentForm.ToggleAddressFilter CStr(CardLabel.Tag)
        Exit Sub
    End If

    If CardLabel.Name = "lblHeaderBlank" Or _
        CardLabel.Name = "lblHeaderBlankRight" Then
        ParentForm.ButtonAction "Reset"
        Exit Sub
    End If

    If Left$(CardLabel.Name, 6) = "photo_" Or _
        Left$(CardLabel.Name, 6) = "phchk_" Or _
        Left$(CardLabel.Name, 6) = "phshd_" Then
        ParentForm.TogglePersonCheck CardLabel.Parent
        Exit Sub
    End If

    If Left$(CardLabel.Name, 4) = "ckr_" Then
        ParentForm.ToggleAttachCheck Mid$(CardLabel.Name, 5)
        Exit Sub
    End If

    If Left$(CardLabel.Name, 3) = "km_" Or Left$(CardLabel.Name, 3) = "ck_" Then
        ParentForm.ToggleAttachCheck Mid$(CardLabel.Name, 4)
        Exit Sub
    End If

    If CardLabel.Name = "lblAttachmentTitle" Or _
        CardLabel.Name = "lblAttachmentAlert" Or _
        CardLabel.Name = "lblAttachmentCount" Then
        ParentForm.ToggleAttachmentFocus
        Exit Sub
    End If

    If Right(CardLabel.Name, 6) = "_title" Or Right(CardLabel.Name, 6) = "_count" Then
        ParentForm.ToggleCategoryFilter CardLabel.Parent
        Exit Sub
    End If

    If Left$(CardLabel.Name, 4) = "row_" Then
        ParentForm.ToggleAttachment Mid$(CardLabel.Name, 5)
        Exit Sub
    End If

    If Left$(CardLabel.Name, 4) = "att_" Then
        ParentForm.ToggleAttachment Mid$(CardLabel.Name, 5)
        Exit Sub
    End If

    If Left$(CardLabel.Name, 3) = "st_" Or _
        Left$(CardLabel.Name, 3) = "pw_" Or _
        Left$(CardLabel.Name, 3) = "mk_" Or _
        Left$(CardLabel.Name, 3) = "fi_" Then
        ParentForm.ToggleAttachment Mid$(CardLabel.Name, 4)
        Exit Sub
    End If

    Set card = CardLabel.Parent
    If TypeName(card) <> "Frame" Then Exit Sub

    If Left$(card.Name, 4) = "att_" Then
        ParentForm.ToggleAttachment Mid$(card.Name, 5)
        Exit Sub
    End If

    ParentForm.ToggleCard card

End Sub
