VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} error_category 
   Caption         =   "Error"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "error_category.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "error_category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub close_button_Click()
'Displays an error message when the
'category field is left blank.

    Unload error_category
End Sub
