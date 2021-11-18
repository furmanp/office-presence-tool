Attribute VB_Name = "Focus_Module"
Option Explicit

Sub Update_Focus(weekNum As Integer)

    Dim focus_wb As workbook
    Dim kow_wb As workbook
    Dim user As String
    Dim column As Variant, row As Variant
    Dim rng_start As Range
    Dim i As Integer
    
    user = Environ("Username")
    Application.ScreenUpdating = False
    Set kow_wb = ThisWorkbook
    Set focus_wb = Workbooks.Open("C:\Users\" + user + "\Pontis\Pontis General - Dokumenty\General\01 Office\Focus.xlsx")

    column = xls_look_(weekNum, vAddress)
    column = Split(column, "$", -1)
    row = xls_look_("PFU - Przemek", vAddress) 'dodac rozpoznawanie imienia
    row = Split(row, "$", -1)
    Set rng_start = Range(column(1) + row(2))
    rng_start.Select

    For i = 0 To 4
        If kow_wb.Sheets("Sender").Cells(3 + i, 3) = "Rv" Then
            rng_start.Offset(0, i) = 1
        End If
    Next i
    
    focus_wb.Close
    Application.ScreenUpdating = True

End Sub

Sub main()
 Update_Focus (46)
End Sub






