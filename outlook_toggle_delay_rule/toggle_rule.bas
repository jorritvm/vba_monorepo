Attribute VB_Name = "toggle_rule"
Sub toggle_wait_rule()
    toggle_rule "send_mail_hold_1_minute"
End Sub


Sub toggle_rule(rule_name As String)
    Dim olRules As Outlook.Rules
    Dim olRule As Outlook.Rule
    Dim rule_status As String
     
    ' fetch the rule
    Set olRules = Application.Session.DefaultStore.GetRules
    Set olRule = olRules.Item(rule_name)
    
    ' toggle
    If olRule.Enabled = True Then
        olRule.Enabled = False
        rule_status = "inactive"
    Else
        olRule.Enabled = True
        rule_status = "active"
    End If
    
    olRules.Save
    
    MsgBox ("Rule '" & rule_name & "' is now set to " & rule_status)
End Sub

