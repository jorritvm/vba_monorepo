VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'--------------------------------------------
' THE FOLLOWING IS RELATED TO THE HTML REPLY
'--------------------------------------------
Private WithEvents oExpl As Explorer
Attribute oExpl.VB_VarHelpID = -1
Private WithEvents oItem As MailItem
Attribute oItem.VB_VarHelpID = -1
Private bDiscardEvents As Boolean
Private olFormat As OlBodyFormat