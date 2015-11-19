Attribute VB_Name = "JobCost"
Option Explicit

'This function returns a markup percantage based off or parts net price
Function Markup(price As Double) As Double
    'Declare variables
    Dim markup_percent As Double
    
    'Initialize variables
    markup_percent = 0
    
    'Logic for determining the markup percent
    If price <= 1 Then
        markup_percent = 5#
    ElseIf price > 1 And price <= 25 Then
        markup_percent = 2.5
    ElseIf price > 25 And price <= 100 Then
        markup_percent = 1.5
    Else
        markup_percent = 1.2
    End If
    
    'Function return
    Markup = markup_percent
End Function

