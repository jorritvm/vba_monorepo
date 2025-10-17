Attribute VB_Name = "appointment_ics"
Sub SendAppointmentAsICS()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olCalendar As Outlook.MAPIFolder
    Dim olItem As Object
    Dim olMail As Outlook.MailItem
    Dim tempICSPath As String
    Dim fso As Object
    Dim ts As Object
    Dim icsContent As String
    
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    
    ' Use the selected item in the calendar
    If TypeName(olApp.ActiveWindow) = "Explorer" Then
        Set olItem = olApp.ActiveExplorer.Selection.Item(1)
    ElseIf TypeName(olApp.ActiveWindow) = "Inspector" Then
        Set olItem = olApp.ActiveInspector.currentitem
    Else
        MsgBox "Please select a calendar item first."
        Exit Sub
    End If
    
    If olItem.Class <> olAppointment Then
        MsgBox "Selected item is not an appointment."
        Exit Sub
    End If
    
    ' Save as ICS
    tempICSPath = Environ("TEMP") & "\" & olItem.Subject & ".ics"
    olItem.SaveAs tempICSPath, olICal
    
    ' Create email
    Set olMail = olApp.CreateItem(olMailItem)
    With olMail
        .To = "jorrit.vm@gmail.com" ' Change to your iCloud address
        .Subject = "Forwarded Appointment: " & olItem.Subject
        .Body = "Attached is an appointment you might want to add to your iPhone calendar."
        .Attachments.Add tempICSPath
        .Send
    End With
    
    ' Clean up
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(tempICSPath) Then fso.DeleteFile tempICSPath
    
    MsgBox "Appointment sent as ICS file."
End Sub
