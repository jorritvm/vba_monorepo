Attribute VB_Name = "attachements"
Sub DeleteAllAttachments()
    '***************************************************************************
    'Purpose: deletes non-embedded attachments from all selected emails in Outlook.
    'Inputs: none
    'Outputs: side effect
    '***************************************************************************
    Dim olkMsg As Object, intIdx As Integer
    Dim userResponse As VbMsgBoxResult

    ' Confirmation prompt
    userResponse = MsgBox("Are you sure you want to delete all non-embedded attachments from the selected emails?", vbYesNo + vbQuestion, "Confirm Deletion")
    If userResponse <> vbYes Then Exit Sub

    ' Proceed with deletion
    For Each olkMsg In Application.ActiveExplorer.Selection
        For intIdx = olkMsg.Attachments.Count To 1 Step -1
            If Not IsHiddenAttachment(olkMsg.Attachments.Item(intIdx)) Then
                olkMsg.Attachments.Item(intIdx).Delete
            End If
        Next
        olkMsg.Close olSave
        Set olkMsg = Nothing
    Next
End Sub
 
Sub SaveAllAttachments()
    '***************************************************************************
    'Purpose: saves non-embedded attachments from selected emails to a folder you choose.
    'Inputs: none
    'Outputs: side effect
    '***************************************************************************
    Const msoFileDialogFolderPicker = 4
    Dim olkMsg As Object, intIdx As Integer, excApp As Object, strPath As String
    Set excApp = CreateObject("Excel.Application")
    With excApp.FileDialog(msoFileDialogFolderPicker)
        .Show
        For intIdx = 1 To .SelectedItems.Count
            strPath = .SelectedItems(intIdx)
        Next
    End With
    If strPath <> "" Then
        For Each olkMsg In Application.ActiveExplorer.Selection
            For intIdx = olkMsg.Attachments.Count To 1 Step -1
                If Not IsHiddenAttachment(olkMsg.Attachments.Item(intIdx)) Then
                    olkMsg.Attachments.Item(intIdx).SaveAsFile strPath & "\" & olkMsg.Attachments.Item(intIdx).FileName
                End If
            Next
            olkMsg.Close olDiscard
            Set olkMsg = Nothing
        Next
    End If
    Set excApp = Nothing
End Sub
 
Private Function IsHiddenAttachment(olkAttachment As Outlook.Attachment) As Boolean
    '***************************************************************************
    'Purpose: Determines if an attachment is embedded.
    '         (like inline images or logos in signatures).
    '***************************************************************************
    Dim olkPA As Outlook.PropertyAccessor
    On Error Resume Next
    Set olkPA = olkAttachment.PropertyAccessor
    IsHiddenAttachment = olkPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7ffe000b")
    On Error GoTo 0
    Set olkPA = Nothing
End Function
