VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} error_description 
   Caption         =   "Error"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "error_description.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "error_description"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub close_button_Click()
'Displays an error message when the description
'field for warranty form is left blank.

    Unload error_description
End Sub
