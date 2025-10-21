Attribute VB_Name = "appointment_ics"
Sub SendAppointmentCopyViaMail()
    '***************************************************************************
    'Purpose: send the selected appointment as ICS mail attachement, modify it so it's unique
    'Inputs: none
    'Outputs: sent email
    '***************************************************************************
    Dim olApp As Outlook.Application
    Dim olItem As Object
    Dim olMail As Outlook.MailItem
    Dim tempICSPath As String
    Dim fso As Object
    Dim ts As Object
    Dim icsLines() As String
    Dim i As Long
    Dim safeSubject As String
    Dim newUID As String
    Dim newDTStamp As String
    Dim SENT_TO_MAIL As String
    
    ' CONFIG
    SENT_TO_MAIL = "me@example.com"
    
    ' Get selected appointment
    Set olApp = Outlook.Application
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
    
    ' Sanitize subject for filename
    safeSubject = Replace(olItem.Subject, ":", "-")
    tempICSPath = Environ("TEMP") & "\" & safeSubject & ".ics"
    
    ' Save original ICS
    olItem.SaveAs tempICSPath, olICal
    
    ' Read and modify ICS content
    newUID = Format(Now, "yyyymmddhhnnss") & "-" & Int(Rnd() * 10000) & "@custom"
    newDTStamp = "DTSTAMP:" & Format(UTCNow(), "yyyymmdd\THhnnss\Z")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(tempICSPath, 1)
    icsLines = Split(ts.ReadAll, vbCrLf)
    ts.Close
    
    For i = LBound(icsLines) To UBound(icsLines)
        If Left(icsLines(i), 7) = "SUMMARY" Then
            icsLines(i) = icsLines(i) & "(Copy)"
        ElseIf Left(icsLines(i), 4) = "UID:" Then
            icsLines(i) = "UID:" & newUID
            ' Remove folded continuation line if present
            If i + 1 <= UBound(icsLines) And (Left(icsLines(i + 1), 1) = " " Or Left(icsLines(i + 1), 1) = vbTab) Then
                icsLines(i + 1) = ""
            End If
        ElseIf Left(icsLines(i), 8) = "DTSTAMP:" Then
            icsLines(i) = newDTStamp
        End If
    Next i
    
    ' Save modified ICS
    Set ts = fso.OpenTextFile(tempICSPath, 2)
    ts.Write Join(icsLines, vbCrLf)
    ts.Close
    
    ' Send email
    Set olMail = olApp.CreateItem(olMailItem)
    With olMail
        .To = SENT_TO_MAIL
        .Subject = "Shared Appointment: " & olItem.Subject
        .Body = "Here's a copy of the appointment with a modified subject."
        .Attachments.Add tempICSPath
        .Send
    End With
    
    ' Clean up
    If fso.FileExists(tempICSPath) Then fso.DeleteFile tempICSPath
    
    MsgBox "ICS with modified subject sent."
End Sub

' ============= HELPERS ================

Function UTCNow() As Date
    UTCNow = Now - (TimeZoneOffset / 24)
End Function

Function TimeZoneOffset() As Double
    ' Adjust this based on your local time zone offset from UTC
    TimeZoneOffset = 2 ' For Brussels during daylight saving time (CEST)
End Function
