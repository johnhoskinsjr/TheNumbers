VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} error_time 
   Caption         =   "Error"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "error_time.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "error_time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub close_error_Click()
'Displays error message when the time
'spent on warranty field was left blank.

    Unload error_time
End Sub
