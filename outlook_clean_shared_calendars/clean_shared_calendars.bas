Attribute VB_Name = "clean_shared_calendars"
'===============================================================
' Module: Shared Calendar Cleanup
'
' Purpose:
'   Automatically clears the "Shared Calendars" list in Outlook
'   when the application is started. This prevents the list of
'   colleagues' calendars from growing too long over time.
'
' Setup Instructions:
'   1. Open Outlook VBA editor (ALT+F11).
'   2. In "ThisOutlookSession", call this method when application quits:
'       Private Sub Application_Startup()
'           clean_shared_calendars_now
'       End Sub
'   3. Restart Outlook. Each time you close Outlook, the
'      shared calendars list will be cleared.
'
' Author: Jorrit Vander Mynsbrugge
' Version: 2025-11-24
'===============================================================


Sub clean_shared_calendars_now()
    Dim calModule As Outlook.CalendarModule
    Dim navPane As Outlook.NavigationPane
    Dim navGroup As Outlook.NavigationGroup
    Dim navFolder As Outlook.NavigationFolder
    Dim foldersToRemove As Collection
    Dim f As Outlook.NavigationFolder
    
    Set navPane = Application.ActiveExplorer.NavigationPane
    Set calModule = navPane.Modules.GetNavigationModule(olModuleCalendar)
    Set foldersToRemove = New Collection
    
    Dim i As Long
    For Each navGroup In calModule.NavigationGroups
        If navGroup.Name = "Shared Calendars" Then
            ' go from back to avoid deleteing from a mutable list and skipping items because of index shift
            For i = navGroup.NavigationFolders.Count To 1 Step -1
                navGroup.NavigationFolders.Remove navGroup.NavigationFolders.Item(i)
            Next
        End If
    Next
End Sub
