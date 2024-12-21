Attribute VB_Name = "activiteit"

'scripthack door JVMY om activiteit in outlook te visualiseren
'2016-02-08
'we tellen in een bepaalde periode aantal items per dag
'we geven een klein rapport in een msgbox en een export in XLS

Public startDate As Date
Public endDate As Date
Public dict As Object
Public totaal As Long

Sub launchpad()
    Dim olApp As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Dim MyFolder As Outlook.MAPIFolder
    Dim msg As String
    Dim excApp As Object, excWkb As Object, excWks As Object
    Dim tick As Long
    
        'initialise
        startDate = CDate(InputBox("What is the startdate for the analysis (format YYYY-m-d)?"))
        endDate = CDate(InputBox("What is the enddate for the analysis (format YYYY-m-d)?"))
        totaal = 0
        tick = 1
        Set dict = CreateObject("Scripting.Dictionary")
        
        'do the recursion
        On Error Resume Next
        Set olApp = Outlook.Application
        Set objNS = olApp.GetNamespace("MAPI")
        Set MyFolder = objNS.PickFolder
        Call ProcessFolder(MyFolder)
        Set objNS = Nothing
     
        'do the export
        Set excApp = CreateObject("Excel.Application")
        With excApp
            .Visible = True
        End With
        Set excWkb = excApp.Workbooks.Add
        Set excWks = excWkb.Worksheets.Add
        excWks.cells(1, 1) = "Datum"
        excWks.cells(1, 2) = "Aantal outlook items"
        For Each o In dict.Keys
            tick = tick + 1
            excWks.cells(tick, 1) = o
            excWks.cells(tick, 2) = dict(o)
        Next
        msg = "Export naar Excel compleet..." & vbCrLf & vbCrLf & " Aantal dagen geanalyseerd: " & CStr(endDate - startDate) & vbCrLf & "Totaal aantal items: " & CStr(totaal) & vbCrLf & "Gemiddeld aantal items per dag: " & CStr(Round(totaal / (endDate - startDate), 2))
        MsgBox msg
    
    Set MyFolder = Nothing
    Set olApp = Nothing
    Set objNS = Nothing
End Sub

 
Sub ProcessFolder(StartFolder As MAPIFolder)
    Dim objFolder As Outlook.MAPIFolder
    Dim objItem As Object
    Dim dateStr As String
    
        On Error Resume Next
        ' process all the items in this folder
        For Each objItem In StartFolder.Items
            Set mai = objItem
            'check if it is in our timewindow
           dateStr = GetDate(mai.ReceivedTime)
            If dateStr >= startDate Then
                If dateStr <= endDate Then
                    'add 1 to the proper dict index
                    If Not dict.Exists(dateStr) Then
                        dict(dateStr) = 0
                    End If
                    dict(dateStr) = CLng(dict(dateStr)) + 1
                    totaal = totaal + 1
                End If
            End If
        Next
        
        ' process all the subfolders of this folder
        For Each objFolder In StartFolder.Folders
            Call ProcessFolder(objFolder)
        Next
     
    'reset references
    Set mai = Nothing
    Set objFolder = Nothing
    Set objItem = Nothing
End Sub

Function GetDate(dt As Date) As String
    GetDate = Year(dt) & "-" & Month(dt) & "-" & Day(dt)
End Function

