VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Warranty_Input 
   Caption         =   "Warranty Request Input"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11535
   OleObjectBlob   =   "Warranty_Input.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Warranty_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This file handles all events for the user form used
'to enter and submit warranty data about jobs.

Option Explicit
Private Sub PopulateCategoryField()
    'Populate category combo box
    With Category
        .AddItem "Kitchen"
        .AddItem "Master Tub"
        .AddItem "Master Shower"
        .AddItem "Lavatory"
        .AddItem "Water Closet"
        .AddItem "Water Heater"
        .AddItem "Leak in Wall"
        .AddItem "Tub/Shower"
        .AddItem "Water Service"
        .AddItem "Outside"
    End With
End Sub

Private Sub close_button_Click()
    
    'Hide warranty data worksheet
    Sheet2.Visible = False
    'Return to Warranty Dashboard
    Sheet3.Activate
    
    'Call subroutine that builds pivot table
    'Call create_pt_warranty
    
    'Close User Form
    Unload Warranty_Input
End Sub

Private Sub SubmitWarrantyButton_Click()
    'Declare variables
    Dim next_row As Long
    Dim time_spent As Double
    Dim hours_spent As Integer
    Dim minutes_spent As Double
    Dim cost As Double
    
    'Initialize variables
    'Find first blank row
    next_row = Application.WorksheetFunction.CountA(Range("A:A")) + 1
    
    'checks hours value if null then assign zero else assign value
    If hours.Text = "" Then
        hours_spent = 0
    Else
        hours_spent = hours.Text
    End If
    'Checks for null on minutes value
    'If null give error else assign value
    If minutes.Text = "" Then
        error_time.Show
        Exit Sub
    Else
        minutes_spent = minutes.Text
    End If
    'Arithmetic for time spent on job
    time_spent = hours_spent + (minutes_spent / 60)
    
    'Check for null on job cost
    If materials_cost = "" Then
        cost = 0
    Else
        cost = materials_cost
    End If
    
    'Check to make sure form is filled out properly
    If Category.Text = "" Then
        error_category.Show
        Exit Sub
    End If
    
    If warranty_desc = "" Then
        error_description.Show
        Exit Sub
    End If
    'Enter information submitted by user form into next empty row
    Cells(next_row, 1) = Cells(next_row - 1, 1) + 1
    Cells(next_row, 2) = Date
    Cells(next_row, 3) = contractor_name.Text
    Cells(next_row, 4) = customer_name.Text
    Cells(next_row, 5) = job_address.Text
    Cells(next_row, 6) = job_city.Text
    Cells(next_row, 7) = Category.Text
    Cells(next_row, 8) = warranty_desc.Text
    Cells(next_row, 9) = time_spent
    Cells(next_row, 10) = cost
    
        
    'Clear Data that was enetered into the user form
    'so user can enter new data.
    Category.Clear
    contractor_name.Text = ""
    customer_name.Text = ""
    job_address.Text = ""
    job_city.Text = ""
    warranty_desc.Text = ""
    materials_cost.Text = ""
    hours.Text = ""
    minutes.Text = ""
    
    'Populate category combo box
    Call PopulateCategoryField
End Sub

Private Sub UserForm_Initialize()
    Call PopulateCategoryField
End Sub

Private Sub UserForm_Terminate()
    'Go back to dashboard
    Sheet3.Activate
    
    'Hide the warranty data worksheet
    Sheet2.Visible = False
    
    'If user form was closed with data not submitted, then open warning
    If contractor_name.Text <> "" Then warning_data_lost.Show
    If customer_name.Text <> "" Then warning_data_lost.Show
    If job_address.Text <> "" Then warning_data_lost.Show
    If job_city.Text <> "" Then warning_data_lost.Show
    If Category.Text <> "" Then warning_data_lost.Show
    If warranty_desc.Text <> "" Then warning_data_lost.Show
    If hours.Text <> "" Then warning_data_lost.Show
    If minutes.Text <> "" Then warning_data_lost.Show
    If materials_cost <> "" Then warning_data_lost.Show
End Sub
