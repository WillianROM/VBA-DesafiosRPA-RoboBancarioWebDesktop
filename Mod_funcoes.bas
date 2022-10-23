Attribute VB_Name = "Mod_funcoes"
Option Explicit

Function WalkEnabledElements(oAutomation As CUIAutomation, element As UIAutomationClient.IUIAutomationElement, strWIndowName As String) As UIAutomationClient.IUIAutomationElement

    Dim walker As UIAutomationClient.IUIAutomationTreeWalker
    
    Set walker = oAutomation.ControlViewWalker
    Set element = walker.GetFirstChildElement(element)
    
    Do While Not element Is Nothing
    
        If InStr(1, element.CurrentName, strWIndowName) > 0 Then
            Set WalkEnabledElements = element
            Exit Function
        End If
        Set element = walker.GetNextSiblingElement(element)
    Loop
End Function

Function PropCondition(UIAutomation As CUIAutomation, Requirement As String, IdType As String) As UIAutomationClient.IUIAutomationCondition
    Select Case IdType
        Case "Name":
            Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_NamePropertyId, Requirement)
        Case "AutoID":
            Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_AutomationIdPropertyId, Requirement)
        Case "ClsName":
            Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_ClassNamePropertyId, Requirement)
        Case "LoczCon":
            Set PropCondition = UIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_LocalizedControlTypePropertyId, Requirement)
    End Select
End Function
