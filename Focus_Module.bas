Attribute VB_Name = "Focus_Module"
Option Explicit
Sub Update_Focus(weekNum As Long)

    Dim focus_wb As Workbook
    Dim kow_wb As Workbook
    Dim user As String, path As String
    Dim column As Variant, row As Variant
    Dim rng_start As Range
    Dim i As Integer
    Dim docs As Variant
    
    docs = Array("Documents", "Documenten", "Dokumenty")
    user = Environ("Username")
    
    Application.ScreenUpdating = False
    Set kow_wb = ThisWorkbook
    
    For i = 0 To UBound(docs)
        On Error Resume Next
        path = "C:\Users\" + user + "\Pontis\Pontis General - " + docs(i) + "\General\01 Office\Focus.xlsx"
        Set focus_wb = Workbooks.Open(path, Editable:=True)
        On Error GoTo 0
    Next i
    
On Error GoTo WrongPath:
    
    focus_wb.Sheets("Office presence").Activate
    column = xls_look_(weekNum, vAddress)
    column = Split(column, "$", -1)
    row = xls_look_(kow_wb.Sheets("Setup").Cells(10, 3), vAddress)
    row = Split(row, "$", -1)
    Set rng_start = Range(column(1) + row(2))
    rng_start.Select

    For i = 0 To 4
        rng_start.Offset(0, i).Clear 'clear previous entries in case of update
        
        If kow_wb.Sheets("Sender").Cells(3 + i, 3) = "Rv" Then 'for Rv see if presence is required or optional
            If kow_wb.Sheets("Sender").Cells(3 + i, 4) = 1 Then
                rng_start.Offset(0, i) = 1
                rng_start.Offset(0, i).Interior.ColorIndex = 6
            ElseIf kow_wb.Sheets("Sender").Cells(3 + i, 4) = 0 Then
                rng_start.Offset(0, i) = 0
                rng_start.Offset(0, i).Interior.ColorIndex = 44
            End If
            rng_start.Offset(0, i).HorizontalAlignment = xlCenter
            
        ElseIf CheckBusyStatus(kow_wb.Sheets("Sender").Cells(3 + i, 3)) Then 'any OFF day highlight with gray
            rng_start.Offset(0, i).Interior.ColorIndex = 44
        End If
    Next i
    
    focus_wb.Save
    focus_wb.Close
    Application.ScreenUpdating = True

WrongPath:
    MsgBox ("FOCUS Spreadhseet could not be found." + vbNewLine + "Synchronize Pontis General")
End Sub






