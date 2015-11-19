VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} warning_data_lost 
   Caption         =   "Warning"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "warning_data_lost.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "warning_data_lost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub close_button_Click()
'Displays a warning messgae when user form is closed
'without submitting data. Warns user data entered will be lost.

    Unload warning_data_lost
End Sub
