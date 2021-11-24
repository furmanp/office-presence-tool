VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Select KOW Options"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    CheckBox_Mail.Value = ActiveWorkbook.Sheets("Setup").Range("E2")
    CheckBox_Calendar.Value = ActiveWorkbook.Sheets("Setup").Range("F2")
    CheckBox_Focus.Value = ActiveWorkbook.Sheets("Setup").Range("G2")

End Sub

Private Sub CommandButton_Save_Click()

    ActiveWorkbook.Sheets("Setup").Range("E2") = CheckBox_Mail.Value
    ActiveWorkbook.Sheets("Setup").Range("F2") = CheckBox_Calendar.Value
    ActiveWorkbook.Sheets("Setup").Range("G2") = CheckBox_Focus.Value

    Unload Me
End Sub

Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub
