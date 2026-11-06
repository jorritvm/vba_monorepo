Attribute VB_Name = "out_of_office"
Private Const OOO_SUBJECT As String = "Jorrit OoO"

Sub CreateOoOAndVACAppointments()
    Dim startDateStr As String
    Dim endDateStr As String
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olCalendar As Outlook.Folder
    Dim olAppt As Outlook.AppointmentItem
    Dim sel As Outlook.Selection
    Dim defaultStart As String
    Dim defaultEnd As String
    Dim v As Outlook.View
    Dim calView As Outlook.CalendarView
    Dim dStart As Date, dEnd As Date

    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olCalendar = olNS.GetDefaultFolder(olFolderCalendar)

    ' Try to get selected appointment date
    Set sel = olApp.ActiveExplorer.Selection
    If sel.Count > 0 Then
        If TypeName(sel.Item(1)) = "AppointmentItem" Then
            defaultStart = Format(sel.Item(1).Start, "yyyy-mm-dd")
            defaultEnd = Format(sel.Item(1).End, "yyyy-mm-dd")
        End If
    End If

    ' If nothing selected, read from the calendar view (works for free-space clicks)
    If defaultStart = "" Then
        Set v = olApp.ActiveExplorer.CurrentView
        If Not v Is Nothing Then
            If TypeName(v) = "CalendarView" Then
                Set calView = v
                dStart = calView.SelectedStartTime
                dEnd = calView.SelectedEndTime
                If dStart <> 0 Then defaultStart = Format(dStart, "yyyy-mm-dd")
                If dEnd <> 0 Then defaultEnd = Format(dEnd, "yyyy-mm-dd")
            End If
        End If
    End If

    ' Final fallback to today if we still don't have a date
    If defaultStart = "" Then defaultStart = Format(Date, "yyyy-mm-dd")
    If defaultEnd = "" Then defaultEnd = defaultStart

    ' Get date inputs, prefill with selected/free-space range
    startDateStr = InputBox("Enter start date (yyyy-mm-dd):", "Start Date", defaultStart)
    If Not IsDate(startDateStr) Then
        MsgBox "Invalid start date format.", vbExclamation
        Exit Sub
    End If

    endDateStr = InputBox("Enter end date (yyyy-mm-dd):", "End Date", defaultEnd)
    If Trim(endDateStr) = "" Then
        endDateStr = startDateStr
    ElseIf Not IsDate(endDateStr) Then
        MsgBox "Invalid end date format.", vbExclamation
        Exit Sub
    End If

    startDate = DateValue(startDateStr)
    endDate = DateValue(endDateStr)

    If endDate < startDate Then
        MsgBox "End date must be after start date.", vbExclamation
        Exit Sub
    End If

    ' Create all-day out-of-office appointment
    Set olAppt = olCalendar.Items.Add(olAppointmentItem)
    With olAppt
        .Subject = OOO_SUBJECT
        .Start = startDate
        .End = endDate + 1
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

