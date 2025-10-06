Attribute VB_Name = "out_of_office"
Sub CreateOoOAndVACAppointments()
    '***************************************************************************
    'Purpose: add out of office to calendar by only setting start- & end-date
    'Inputs: none
    'Outputs: side effect
    '***************************************************************************
    Dim startDateStr As String
    Dim endDateStr As String
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olCalendar As Outlook.Folder
    Dim olAppt As Outlook.AppointmentItem

    ' Get date inputs
    startDateStr = InputBox("Enter start date (yyyy-mm-dd):", "Start Date")
    If Not IsDate(startDateStr) Then
        MsgBox "Invalid start date format.", vbExclamation
        Exit Sub
    End If

    endDateStr = InputBox("Enter end date (yyyy-mm-dd):", "End Date")
    If Trim(endDateStr) = "" Then
        endDateStr = startDateStr
    ElseIf Not IsDate(endDateStr) Then
        MsgBox "Invalid end date format.", vbExclamation
        Exit Sub
    Else
        endDate = DateValue(endDateStr)
    End If

    startDate = DateValue(startDateStr)
    endDate = DateValue(endDateStr)

    If endDate < startDate Then
        MsgBox "End date must be after start date.", vbExclamation
        Exit Sub
    End If

    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olCalendar = olNS.GetDefaultFolder(olFolderCalendar)

    ' Create all-day "Jorrit OoO" appointment
    Set olAppt = olCalendar.Items.Add(olAppointmentItem)
    With olAppt
        .Subject = "Jorrit OoO"
        .Start = startDate
        .End = endDate + 1 ' All-day appointments end the next day
        .AllDayEvent = True
        .BusyStatus = olOutOfOffice
        .Save
    End With

    ' Create daily "VAC" appointments
    currentDate = startDate
    Do While currentDate <= endDate
        Set olAppt = olCalendar.Items.Add(olAppointmentItem)
        With olAppt
            .Subject = "VAC"
            .Start = currentDate + TimeValue("07:00:00")
            .End = currentDate + TimeValue("18:00:00")
            .BusyStatus = olOutOfOffice
            .Save
        End With
        currentDate = currentDate + 1
    Loop

    MsgBox "Appointments created successfully!", vbInformation
End Sub

