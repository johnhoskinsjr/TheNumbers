Attribute VB_Name = "ProjectionOfIncome"
Option Explicit

Function PaidSoFar(paid_date As Range, Amount As Range, invoice_date As Range) As Double
'Returns the sum of invoices paid currenty according
'to invoice sent date of current row. Only sums invoices
'paid in the current month, and resets to zero each month.
'This function is meant to be a column formula in a table.

    'Declare variables
    Dim cell As Range
    Dim paid_amount As Double
    
    'Loop through invoice date range
    For Each cell In paid_date
        'paid date is less than the invoice date, invoice is already paid
        'and invoice paid in same month of invoice sent
        If DateValue(invoice_date) >= DateValue(cell) _
            And month(DateValue(invoice_date)) = month(DateValue(cell)) Then
            'Sum up the invoices
             paid_amount = paid_amount + (Amount(cell.Row - 1))
             'Debug.Print (Amount(cell.Row - 1))
        End If
    Next cell
    'Function return
    PaidSoFar = paid_amount
End Function
