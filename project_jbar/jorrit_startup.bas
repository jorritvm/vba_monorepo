Attribute VB_Name = "jorrit_startup"

'wanneer draait Project_Open???
'https://msdn.microsoft.com/en-us/library/bb648784(v=office.12).aspx
'EPRM: kopieer onderstaande code per project package


Private Sub Project_Open(ByVal pj As Project)
  On Error Resume Next
  ' users zetten hier hun triggers
    Application.Macro "global.MPT!jorrit_functional.createjbar"
End Sub


Private Sub Project_BeforeClose(ByVal pj As Project)
  On Error Resume Next
  ' users zetten hier hun triggers
  Application.Macro "global.MPT!jorrit_functional.deletejbar"
End Sub


