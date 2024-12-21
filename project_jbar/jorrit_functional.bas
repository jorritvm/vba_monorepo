Attribute VB_Name = "jorrit_functional"

Public myOLApp As Object
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
Dim Relationshipdiagram_navigator_stop As Boolean


Sub createjbar()
    Dim cbar As CommandBar
    Dim cctrl As CommandBarButton
           
    deletejbar

    Set cbar = CommandBars.Add("JBar")
    cbar.Visible = True
    cbar.Position = msoBarTop
    cbar.Left = CommandBars("Standard").Left + CommandBars("Standard").Width
    cbar.RowIndex = CommandBars("Standard").RowIndex
                   
    With cbar
    'automatic calculation
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .BeginGroup = True
            .State = msoButtonUp
            .FaceId = 80 'change for icon
            .OnAction = "Macro ""calcauto"""
        End With
    'manual calculation
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .State = msoButtonUp
            .FaceId = 92
            .OnAction = "Macro ""calcman"""
        End With
    'Activate_Relationshipdiagram_navigator
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .BeginGroup = True
            .State = msoButtonUp
            .FaceId = 88
            .OnAction = "Macro ""Activate_Relationshipdiagram_navigator"""
        End With
    'Deactivate_Relationshipdiagram_navigator
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .State = msoButtonUp
            .FaceId = 94
            .OnAction = "Macro ""Deactivate_Relationshipdiagram_navigator"""
        End With
    'reset DATE format
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .BeginGroup = True
            .State = msoButtonUp
            .FaceId = 33
            .OnAction = "Macro ""resetdateformat"""
        End With
    'white
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .BeginGroup = True
            .Caption = "White"
            .State = msoButtonUp
            .Style = msoButtonIcon 'change for icon
            .OnAction = "Macro ""color_white"""
            .Picture = stdole.StdFunctions.LoadPicture(Environ("USERPROFILE") & "\msprojectkleuren\white.bmp")
        End With
         ' red
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .Caption = "Red"
            .State = msoButtonUp
            .Style = msoButtonIcon 'change for icon
            .OnAction = "Macro ""color_red"""
            .Picture = stdole.StdFunctions.LoadPicture(Environ("USERPROFILE") & "\msprojectkleuren\red.bmp")
        End With
        'yellow
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .Caption = "Yellow"
            .State = msoButtonUp
            .Style = msoButtonIcon 'change for icon
            .OnAction = "Macro ""color_yellow"""
            .Picture = stdole.StdFunctions.LoadPicture(Environ("USERPROFILE") & "\msprojectkleuren\yellow.bmp")
        End With
        'green
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .Caption = "Green"
            .State = msoButtonUp
            .Style = msoButtonIcon 'change for icon
            .OnAction = "Macro ""color_green"""
            .Picture = stdole.StdFunctions.LoadPicture(Environ("USERPROFILE") & "\msprojectkleuren\green.bmp")
        End With
        'blue
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .Caption = "Blue"
            .State = msoButtonUp
            .Style = msoButtonIcon 'change for icon
            .OnAction = "Macro ""color_blue"""
            .Picture = stdole.StdFunctions.LoadPicture(Environ("USERPROFILE") & "\msprojectkleuren\blue.bmp")
        End With
        'Darkblue
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .Caption = "Darkblue"
            .State = msoButtonUp
            .Style = msoButtonIcon 'change for icon
            .OnAction = "Macro ""color_darkblue"""
            .Picture = stdole.StdFunctions.LoadPicture(Environ("USERPROFILE") & "\msprojectkleuren\darkblue.bmp")
        End With
        'fushia
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .Caption = "Fushia"
            .State = msoButtonUp
            .Style = msoButtonIcon 'change for icon
            .OnAction = "Macro ""color_fushia"""
            .Picture = stdole.StdFunctions.LoadPicture(Environ("USERPROFILE") & "\msprojectkleuren\fushia.bmp")
        End With
        'grey
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .Caption = "Grey"
            .State = msoButtonUp
            .Style = msoButtonIcon 'change for icon
            .OnAction = "Macro ""color_grey"""
            .Picture = stdole.StdFunctions.LoadPicture(Environ("USERPROFILE") & "\msprojectkleuren\grey.bmp")
        End With
        'export outlook calendar
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .BeginGroup = True
            .Caption = "Export to outlook calendar"
            .State = msoButtonUp
            .FaceId = 2473 'change for icon
            .OnAction = "Macro """ & "Export_Selection_To_OL_Appointments"""
        End With
        'export outlook task
        Set ctrl = .Controls.Add(msoControlButton, , , , True)
        With ctrl
            .Caption = "Export to outlook task"
            .State = msoButtonUp
            .FaceId = 837 'change for icon
            .OnAction = "Macro """ & "Export_Selection_To_OL_Tasks"""
        End With
    End With
    Set ctrl = Nothing
End Sub

Sub deletejbar()
    On Error Resume Next
    CommandBars("JBar").Delete
End Sub

Sub calcauto()
    Application.Calculation = pjAutomatic
End Sub

Sub calcman()
    Application.Calculation = pjManual
End Sub

Sub resetdateformat()
    Application.DefaultDateFormat = pjDate_ddd_mm_dd_yy
End Sub

'set colors
'1 red
'2 yellow
'3 green
'4 blue
'5 dark blue
'6 fushia
'7 white
Sub color_x(x)
    Set ts = ActiveSelection.Tasks
    For Each tsk In ts
        FontEx CellColor:=x
    Next tsk
End Sub
Sub color_red()
    Call color_x(1)
End Sub
Sub color_yellow()
    Call color_x(2)
End Sub
Sub color_green()
    Call color_x(3)
End Sub
Sub color_blue()
    Call color_x(4)
End Sub
Sub color_darkblue()
    Call color_x(5)
End Sub
Sub color_fushia()
    Call color_x(6)
End Sub
Sub color_white()
    Call color_x(7)
End Sub
Sub color_grey()
    Call color_x(15)
End Sub

Sub Export_Selection_To_OL_Appointments()
    Dim myTask As Task
    Dim myItem As Object
    Dim tick As Integer, iReply As Integer
      
    On Error Resume Next
    Set myOLApp = CreateObject("Outlook.Application")
    tick = 0
    
    For Each myTask In ActiveSelection.Tasks
      tick = tick + 1
    Next myTask
    
    iReply = MsgBox(Prompt:="Are you sure you want to export " & CStr(tick) & " appointment(s)?", Buttons:=vbYesNoCancel, Title:="Confirmation")
    If iReply = vbYes Then
        For Each myTask In ActiveSelection.Tasks
            Set myItem = myOLApp.CreateItem(1)
            With myItem
              .Start = myTask.Start
              .End = myTask.Finish
              .Subject = myTask.Name & " (EPRM)"
              .Categories = myTask.Project
              .Body = myTask.Notes
              .BusyStatus = 0
              .AllDayEvent = True
              .Save
            End With
        Next myTask
        MsgBox "Successfully exported " & CStr(tick) & " appointment(s)."
    ElseIf iReply = vbNo Then
       'Do Other Stuff
    Else 'They cancelled (VbCancel)
        Exit Sub
    End If
    
End Sub

Sub Export_Selection_To_OL_Tasks()
Dim myTask As Task
Dim myItem As Object
   
  On Error Resume Next
  Set myOLApp = CreateObject("Outlook.Application")
  
  For Each myTask In ActiveSelection.Tasks
    Set myItem = myOLApp.CreateItem(3)
    With myItem
      .StartDate = myTask.Start
      .DueDate = myTask.Finish
      .Subject = myTask.Name & " (EPRM)"
      .Body = myTask.Notes
      .Categories = myTask.Project
      .Save
    End With
  Next myTask

End Sub

Sub Activate_Relationshipdiagram_navigator()
    Relationshipdiagram_navigator_stop = False
    ActiveWindow.TopPane.Activate
    Application.OutlineShowAllTasks
    oldTaskID = ActiveSelection.Tasks(1).UniqueID
    newTaskID = ActiveSelection.Tasks(1).UniqueID
    Do
        If ActiveWindow.ActivePane.View.Name = "Relationship Diagram" Then
            newTaskID = ActiveSelection.Tasks(1).UniqueID
            taskrow = ActiveSelection.Tasks(1).ID
            If newTaskID <> oldTaskID Then 'selection has changed
                ActiveWindow.TopPane.Activate
                Application.Find Field:="UniqueID", test:="equals", Value:=newTaskID
                SelectTaskCell Row:=taskrow, Column:="Name", RowRelative:=False
                'Sleep (500)
            End If
            oldTaskID = newTaskID
        End If
        DoEvents
    Loop Until Relationshipdiagram_navigator_stop
End Sub

Sub Deactivate_Relationshipdiagram_navigator()
    Relationshipdiagram_navigator_stop = True
End Sub


