Attribute VB_Name = "html_reply"
'--------------------------------------------
' THE FOLLOWING IS RELATED TO THE HTML REPLY
'--------------------------------------------
Private Sub oExpl_SelectionChange()
   On Error Resume Next
   Set oItem = oExpl.Selection.item(1)
End Sub

Private Sub oItem_Reply(ByVal Response As Object, Cancel As Boolean)
    If bDiscardEvents Or oItem.BodyFormat = olFormat Then
       Exit Sub
    End If
    Cancel = True
    bDiscardEvents = True
    Dim oResponse As MailItem
    Set oResponse = oItem.Reply
    oResponse.BodyFormat = olFormat
    oResponse.Display
    bDiscardEvents = False
End Sub

Private Sub oItem_ReplyAll(ByVal Response As Object, Cancel As Boolean)
   If bDiscardEvents Or oItem.BodyFormat = olFormat Then
       Exit Sub
   End If
   Cancel = True
   bDiscardEvents = True
  Dim oResponse As MailItem
   Set oResponse = oItem.ReplyAll
   oResponse.BodyFormat = olFormat
   oResponse.Display
   bDiscardEvents = False
End Sub

Private Sub oItem_Forward(ByVal Forward As Object, Cancel As Boolean)
   If bDiscardEvents Or oItem.BodyFormat = olFormat Then
       Exit Sub
   End If
   Cancel = True
   bDiscardEvents = True
   Dim oResponse As MailItem
   Set oResponse = oItem.Forward
   oResponse.BodyFormat = olFormat
   oResponse.Display
   bDiscardEvents = False
End Sub







