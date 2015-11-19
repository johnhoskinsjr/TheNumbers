VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} delete_warranty_form 
   Caption         =   "Delete Warranty"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7665
   OleObjectBlob   =   "delete_warranty_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "delete_warranty_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub close_button_Click()
    'Close the delete warranty form
    Unload delete_warranty_form
    
End Sub

Private Sub delete_button_Click()
    'Assign last row in table to variable
    Dim last_row As Integer
    Dim i As Integer
    last_row = Application.WorksheetFunction.CountA(Range("A:A"))
    
    'Assign selected row to variable
    Dim list_index As Integer
    list_index = warranty_list.ListIndex
    
    'Delete the selected item from list
    'Debug.Print (last_row - list_index)
    Rows(last_row - list_index).Delete
    
    'Clear list box before adding items
    warranty_list.Clear
    
    'Redeclare row length
    last_row = Application.WorksheetFunction.CountA(Range("A:A"))
    'Load last 20 warranty request into list
    For i = 1 To 20
        warranty_list.AddItem
        warranty_list.List(i - 1, 0) = Cells(last_row, 2)
        warranty_list.List(i - 1, 1) = Cells(last_row, 5)
        warranty_list.List(i - 1, 2) = Cells(last_row, 7)
'        Debug.Print (last_row)
        last_row = last_row - 1
        If last_row = 1 Then Exit For
    Next i
    
End Sub

Private Sub UserForm_Initialize()
    'Make active worksheet warranty data
    Sheet2.Visible = True
    Sheet2.Activate
    
    'Find starting row
    Dim last_row As Integer
    Dim i As Integer
    last_row = Application.WorksheetFunction.CountA(Range("A:A"))
    
    'Load last 20 warranty request into list
    For i = 1 To 20
        warranty_list.AddItem
        warranty_list.List(i - 1, 0) = Cells(last_row, 2)
        warranty_list.List(i - 1, 1) = Cells(last_row, 5)
        warranty_list.List(i - 1, 2) = Cells(last_row, 7)
'        Debug.Print (last_row)
        last_row = last_row - 1
        If last_row = 1 Then Exit For
    Next i
End Sub

Private Sub UserForm_Terminate()
    'Make active sheet dashboard and hide warranty data
    Sheet3.Activate
    Sheet2.Visible = False
    
    'Clear warranty list box
    warranty_list.Clear
    
End Sub
