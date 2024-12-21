VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' THE FOLLOWING IS RELATED TO THE HTML REPLY MACRO
Private WithEvents oExpl As Explorer
Attribute oExpl.VB_VarHelpID = -1
Private WithEvents oItem As MailItem
Attribute oItem.VB_VarHelpID = -1
Private bDiscardEvents As Boolean
Private olFormat As OlBodyFormat

' THE FOLLOWING IS RELATED TO THE MAILSORTER
Dim WithEvents colAppItems As Outlook.Items
Attribute colAppItems.VB_VarHelpID = -1
Dim WithEvents colSentItems As Items
Attribute colSentItems.VB_VarHelpID = -1


Private Sub Application_Startup()
    ' THE FOLLOWING IS RELATED TO THE MAILSORTER
    Set colAppItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
    Set colSentItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderSentMail).Items
    
    'THE FOLLOWING IS RELATED TO THE HTML REPLY MACRO
    'Set oExpl = Application.ActiveExplorer
    'bDiscardEvents = False
    'olFormat = olFormatPlain        '(*1) - reply using plain text
    'olFormat = olFormatHTML        '(*2) - reply using HTML
End Sub

'--------------------------------------------
' THE FOLLOWING IS RELATED TO THE HTML REPLY
'--------------------------------------------
Private Sub oExpl_SelectionChange()
   On Error Resume Next
   Set oItem = oExpl.Selection.Item(1)
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


'--------------------------------------------
' THE FOLLOWING IS RELATED TO THE MAILSORTER
'--------------------------------------------
'rev1.2: works with appointments as well
'rev1.1: fixed crash when no correct value given; now shows title in messagebox
'rev1.0: itemadd instead of item_send for compatibility with multiple mailboxes
'rev0.x: works for one mailbox, for manually initiated sort and on-send sort

Public Function afkorting(tekst As String) As MAPIFolder
    'vertaalt afkorting naar MAPIFolder object
    Dim objFolder As Outlook.MAPIFolder
    Select Case tekst
        'algemeen
        'Case "gen"
        '    Set objFolder = GetFolder("Personal Folders\" + Format(Date, "yyyy-mm"))
        'Case "gen-"
        '    Set objFolder = GetFolder("Personal Folders\" + Format(DateAdd("m", -1, Date), "yyyy-mm"))
        Case "gen"
            Set objFolder = GetFolder("Personal Folders\" + Format(Date, "yyyy"))
        Case "opl"
            Set objFolder = GetFolder("Personal Folders\archief\01. Elia\opleidingen")
        Case "car"
            Set objFolder = GetFolder("Personal Folders\archief\01. Elia\leasing")
        Case "it"
            Set objFolder = GetFolder("Personal Folders\archief\01. Elia\IT tips")
        Case "obj"
            Set objFolder = GetFolder("Personal Folders\archief\10. Objectieven\algemeen")
        Case "priv"
            Set objFolder = GetFolder("Personal Folders\archief\z_prive")
        Case "fun"
            Set objFolder = GetFolder("Personal Folders\archief\z_fun")

        
        'sma
        Case "gd"
            Set objFolder = GetFolder("Personal Folders\archief\04. SMA\team")
        Case "mst"
            Set objFolder = GetFolder("Personal Folders\archief\04. SMA\tyndp_mst")
        Case "srv"
            Set objFolder = GetFolder("Personal Folders\archief\04. SMA\SRV6")
        Case "fop"
            Set objFolder = GetFolder("Personal Folders\archief\04. SMA\FOP")
            
        'ipm
        Case "ipm"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\01. IPM")
        Case "em"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\40. REX - EM coreteam")
        Case "team"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\25. Team")
        'projecten
        Case "lil"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\20. Projecten\TR40348 Lillo 380 + 150")
        Case "rui"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\20. Projecten\TR40720 Ruien 150")
        Case "kno"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\20. Projecten\TR42045 Knokke")
        Case "ny"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\20. Projecten\TR42137 Bruny")
        Case "dro"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\20. Projecten\TR40885 Drogenbos")
        Case "ave"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\20. Projecten\TR42230 Avelgem")
        Case "ron"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\20. Projecten\TR42240 Ronse")
        Case "moe"
            Set objFolder = GetFolder("Personal Folders\archief\z_old\03. EE\20. Projecten\TR42327 Moeskroen ERS + 150")
    End Select
    Set afkorting = objFolder
End Function

'Count
'Attribute VB_Name = "CountSelected"
Sub CountSelectedItems()
    Dim objSelection As Outlook.Selection
    Dim Result As Integer
    Set objSelection = Application.ActiveExplorer.Selection
    Result = MsgBox("Number of selected items: " & _
                objSelection.Count, vbInformation, "Selected Items")
End Sub

Public Sub colSentItems_ItemAdd(ByVal Item As Object)
  'If Item.Class = olMail Then
    'Dim objItem As Outlook.MailItem
    Set objItem = Item
   Set objFolder = ask(objItem)
    'Set Item.SaveSentMessageFolder = objFolder 'werkt enkel als we in dezelfde mailbox blijven
    If TypeName(objFolder) <> "Nothing" Then
        Item.Move objFolder
    End If
  'End If
End Sub

Private Sub colAppItems_ItemAdd(ByVal Item As Object)
  On Error Resume Next
  Dim Appt As Outlook.AppointmentItem
  If TypeOf Item Is Outlook.AppointmentItem And Item.Start < Now() Then
    Set Appt = Item
     Appt.ReminderSet = False
     Appt.Save
  End If
End Sub

Sub move_item()
    Dim objFolder As MAPIFolder
    'vragen naar project
    Set objFolder = ask(Nothing)
    If TypeName(objFolder) <> "Nothing" Then
        Call MoveSelectedMessagesToFolder(objFolder)
    End If
End Sub

Public Function ask(objItem) As MAPIFolder
    Dim objFolder As MAPIFolder
    Dim objNS As NameSpace
    Set objNS = Application.GetNamespace("MAPI")
    
    If Not objItem Is Nothing Then
        tekst = InputBox("Enter map shortcut for mail:" + vbNewLine + objItem.Subject)
    Else:
        tekst = InputBox("Map shortcut?")
    End If
    If tekst <> "" Then
        'gebruik van afkorting
        Set objFolder = afkorting(CStr(tekst))
    End If
    'dan via tree
    If TypeName(objFolder) = "Nothing" Then
        Set objFolder = objNS.PickFolder
    End If
        
    'als het nog niet gekozen is is het een CANCEL
    If TypeName(objFolder) = "Nothing" Then
        'Set objNS = Application.GetNamespace("MAPI")
        'Set objFolder = objNS.GetDefaultFolder(olFolderInbox)
    End If
    Set ask = objFolder
End Function

Public Function GetFolder(strFolderPath As String) As MAPIFolder
  ' used by afkorting
  ' strFolderPath needs to be something like
  '   "Public Folders\All Public Folders\Company\Sales" or
  '   "Personal Folders\Inbox\My Folder"
  Dim objApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Dim colFolders As Outlook.Folders
  Dim objFolder As Outlook.MAPIFolder
  Dim arrFolders() As String
  Dim I As Long
  On Error Resume Next

  strFolderPath = Replace(strFolderPath, "/", "\")
  arrFolders() = Split(strFolderPath, "\")
  Set objApp = Application
  Set objNS = objApp.GetNamespace("MAPI")
  Set objFolder = objNS.GetDefaultFolder(olFolderInbox).Parent
  Set objFolder = objFolder.Folders.Item(arrFolders(0))
  If Not objFolder Is Nothing Then
    For I = 1 To UBound(arrFolders)
      Set colFolders = objFolder.Folders
      Set objFolder = Nothing
      Set objFolder = colFolders.Item(arrFolders(I))
      If objFolder Is Nothing Then
        Exit For
      End If
    Next
  End If

  Set GetFolder = objFolder
  Set colFolders = Nothing
  Set objNS = Nothing
  Set objApp = Nothing
End Function

Sub MoveSelectedMessagesToFolder(Folder As MAPIFolder)
'On Error Resume Next
    Dim objFolder As Outlook.MAPIFolder
    'Dim objItem As Outlook.AppointmentItem
    Set objFolder = Folder

    If objFolder Is Nothing Then
        MsgBox "This folder doesn't exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
        Exit Sub
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        'Require that this procedure be called only when a message is selected
        Exit Sub
    End If

    For Each objItem In Application.ActiveExplorer.Selection
        objItem.UnRead = False
        objItem.Move objFolder
    Next

    Set objItem = Nothing
    Set objFolder = Nothing
End Sub


