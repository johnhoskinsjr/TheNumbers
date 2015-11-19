Attribute VB_Name = "MonthlyIncome"
Option Explicit

Function MonthlyIncome(date_paid As Range, Amount As Range, invoice_date As Range) As Double
'This function sums the total invoices paid for the
'current month. This is a formula meant for a column
'in a table. The data output will be used for values
'in other column formulas.

    'Declare variables
    Dim income As Double
    Dim month_wanted As Integer
    
    'Initialize variables
    month_wanted = month(DateValue(invoice_date.Value))
    
    'Set income variable
    income = ActualPaid(date_paid, Amount, month_wanted)
    
    'Function return
    MonthlyIncome = income
End Function

Function ProjectFirstMonth(invoice_date As Date, _
    invoice_amount As Currency, age_rng As Range) As Currency
'This formula makes projection of amount that should be
'paid in for month for the invoice of this row.

    'Declare variables
    Dim age As Integer
    Dim last_day_month As Date
    Dim prob_paid As Double
    
    'Initialize variable
    last_day_month = LastDayOfMonth(invoice_date) 'Set value of last day of month for the current month
    age = last_day_month - invoice_date 'Sets the age of invoice based off the difference of date invoice sent and end of month
    'Sets probability of being paid this month
    prob_paid = Application.WorksheetFunction.CountIf(age_rng, "<=" & age) / Application.WorksheetFunction.Count(age_rng)

    'Function return
    ProjectFirstMonth = invoice_amount * prob_paid
    
End Function

Function LastDayOfMonth(current_date As Date, Optional add_month As Integer = 0) As Date
'A function that takes input of a date and returns
'the last day of the month.

    'Declare variables
    Dim last_day As Date
    Dim month_date As Integer
    
    'Initialize variables
    month_date = month(DateValue(current_date)) + add_month
    
    'Case statement that assigns last day of month to last_day
    'Based off the month number of date passed as aurgument
    Select Case month_date
        Case 1
            'If january
            last_day = #1/31/2015#
        Case 2
            'If febuary
            last_day = #2/28/2015#
        Case 3
            'If march
            last_day = #3/31/2015#
        Case 4
            'If april
            last_day = #4/30/2015#
        Case 5
            'If may
            last_day = #5/31/2015#
        Case 6
            'If june
            last_day = #6/30/2015#
        Case 7
            'If july
            last_day = #7/31/2015#
        Case 8
            'If august
            last_day = #8/31/2015#
        Case 9
            'If september
            last_day = #9/30/2015#
        Case 10
            'If october
            last_day = #10/31/2015#
        Case 11
            'If november
            last_day = #11/30/2015#
        Case 12
            'If december
            last_day = #12/31/2015#
    End Select
    'Function Return
    LastDayOfMonth = last_day
End Function

Function ProjectSecondMonth(invoice_date As Date, invoice_amount As Currency, _
    invoice_age As Range, first_proj As Currency) As Currency
'This formula makes projection of amount that should be
'paid in for next month for the invoice of this row.

    'Declare variables
    Dim first_day As Date   'Variable that hold first day of month
    Dim last_day As Date    'Variable that holds last day of the month
    Dim early_age As Integer    'Variable that hold invoice age for first day of month
    Dim old_age As Integer  'Variable that holds invoice age for last day of month
    Dim prob_paid As Double 'Variable that holds the probability of begin paid based on invoice date
    Dim difference_amount As Currency
    Dim second_project As Currency
    Dim actual_project As Currency
    
    'Initialize variables
    'Assign dates to variables
    first_day = FindFirstDay(invoice_date, 1)
    last_day = LastDayOfMonth(invoice_date, 1)
    
    'Determine the age based off date
    early_age = first_day - invoice_date - 1
    old_age = last_day - invoice_date - 1
    
    'Determine the probability of being paid for invoice based on age
    prob_paid = Application.WorksheetFunction.CountIfs(invoice_age, ">=" & early_age, invoice_age, "<=" & old_age) _
        / Application.WorksheetFunction.Count(invoice_age)
    
    'Set project variable for double check of over projection
    second_project = invoice_amount * prob_paid
    difference_amount = invoice_amount - first_proj
    
    'Make sure second month projection doesnt over project the amount of invoice
    If second_project > difference_amount Then
        actual_project = difference_amount
    Else
        actual_project = second_project
    End If
    'Function return
    ProjectSecondMonth = actual_project
End Function

Function FindFirstDay(current_date As Date, Optional add_month As Integer = 0) As Date
'This functions takes a date as input and returns
'the first day of the month for that date.

    'Declare variables
    Dim first_day As Date
    Dim month_date As Integer
    
    'Initialize variables
    month_date = month(current_date) + add_month
    If month_date = 13 Then month_date = 1
    
    'Case statement that switches based off month number
    'assigns the first day of the month
     Select Case month_date
        Case 1
            'If january
            first_day = #1/1/2016#
        Case 2
            'If febuary
            first_day = #2/1/2015#
        Case 3
            'If march
            first_day = #3/1/2015#
        Case 4
            'If april
            first_day = #4/1/2015#
        Case 5
            'If may
            first_day = #5/1/2015#
        Case 6
            'If june
            first_day = #6/1/2015#
        Case 7
            'If july
            first_day = #7/1/2015#
        Case 8
            'If august
            first_day = #8/1/2015#
        Case 9
            'If september
            first_day = #9/1/2015#
        Case 10
            'If october
            first_day = #10/1/2015#
        Case 11
            'If november
            first_day = #11/1/2015#
        Case 12
            'If december
            first_day = #12/1/2015#
    End Select
    
    'Function Return
    FindFirstDay = first_day
End Function

Function ProjectThirdMonth(invoice_date As Date) As Integer
'This formula makes projection of amount that should be
'paid for 2 months for the invoice of this row.

    'Declare variables
    Dim first_day As Date   'Variable that hold first day of month
    Dim last_day As Date    'Variable that holds last day of the month
    Dim early_age As Integer    'Variable that hold invoice age for first day of month
    Dim old_age As Integer  'Variable that holds invoice age for last day of month
    
    'Initialize variables
    'Assign dates to variables
    first_day = FindFirstDay(invoice_date, 2)
    last_day = LastDayOfMonth(invoice_date, 2)
    
    'Determine the age based off date
    early_age = first_day - invoice_date - 1
    old_age = last_day - invoice_date - 1
    
    'Function return
    ProjectThirdMonth = old_age
    
End Function

Function IncludeFirstProject(current_date As Date, invoice_date As Range) As Currency
'This function makes a total prediction of the amount
'expected to receive payment for the current month.
'This return the total currency amount expected for
'current date. This formula should be used with a column in table.

    'Declare Variables
    Dim open_invoices As Integer
    Dim paid_invoices As Integer
    Dim row_number As Integer
    Dim cell As Range
    Dim current_projection As Currency
    Dim ws As Worksheet
    
    'Initialize variables
    Set ws = Worksheets("income_data")
    'open_invoices = Application.WorksheetFunction.CountIf(invoice_date, "<=" & current_date)
    'paid_invoices = Application.WorksheetFunction.CountIf(paid_date, "<=" & current_date)
    For Each cell In invoice_date
        'If statement that equal to uppaid invoices
        'If invoices sent out before current date
        'and invoices who were paid after current date
        If cell <= current_date And current_date <= ws.Cells(cell.Row, 4) Then
            'Sums up projected income of first month for unpaid invoices
            current_projection = current_projection + ws.Cells(cell.Row, 7)
            'If open invoice is a month old, then add second month project for that invoice
            If month(current_date) > month(ws.Cells(cell.Row, 4)) Then
                current_projection = current_projection + ws.Cells(cell.Row, 8)
            End If
            'Debug.Print (current_projection)
        End If
    Next cell
    
    'Function return
    IncludeFirstProject = current_projection
    
End Function

Function SumOpenInvoices(current_date As Date, paid_rng As Range) As Currency
'This function returns the open of all current open invoices.
'This function should be used with a column in a table, and
'used as a value in another function.

    'Declare variables
    Dim total As Currency
    Dim cell As Range
    Dim ws As Worksheet
    
    'Initialize variables
    Set ws = Worksheets("income_data")
    'Loop through paid invoices range
    For Each cell In paid_rng
        'Check to see if date paid is after current date
        'and that invoice was sent out before current date
        If cell >= current_date And current_date >= ws.Cells(cell.Row, 3) Then
            total = total + ws.Cells(cell.Row, 2)
        End If
    Next cell
    'Function return
    SumOpenInvoices = total
End Function
