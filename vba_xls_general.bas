Attribute VB_Name = "vba_xls_general"
Option Explicit

Public Enum vOutput
    vAddress
    vValue
    vWidth
    vLength
End Enum


'###################################################################
'   16.12.19 xls_look
'
'-------------------------------------------------------------------
Function xls_look_(sRef As Variant, output As vOutput, Optional offsetDown As Long = 0, Optional offsetRight As Long = 0)

    xls_look_ = "Not Found"                     'default result
    Dim sTxt As String
    Dim rCell As Range
   
    '   --------------------------check sRef is a cell reference-----------------------------------------
    On Error Resume Next
        Set rCell = Cells.Find(sRef, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        sTxt = rCell.Offset(offsetDown, offsetRight).Address
    If Err.Number > 0 Then Exit Function
    
    Select Case output
        Case vAddress:
            xls_look_ = Range(sTxt).Address
        Case vValue:
            xls_look_ = Range(sTxt).Value
        Case vLength:
            If Range(sTxt).Offset(1, 0).Value <> "" Then
                xls_look_ = Range(Range(sTxt), Range(sTxt).End(xlDown)).Count
            Else
                xls_look_ = 1
            End If
        Case vWidth:
            If Range(sTxt).Offset(0, 1).Value <> "" Then
                xls_look_ = Range(Range(sTxt), Range(sTxt).End(xlToRight)).Count
            Else
                xls_look_ = 1
            End If
    End Select
End Function

'###################################################################
'   15.01.26    adds data into an array
'   sRef: text to find
'   can set n_down_end and n_right_end = 0 and will automatically find end positions
'-------------------------------------------------------------------
Function a_make(sRef As String, Optional n_down_start As Long = 0, Optional n_right_start As Long = 0, Optional n_down_end As Long = 0, Optional n_right_end As Long = 0)

    Dim rCell As Range
    
    On Error Resume Next
    Set rCell = Range(xls_look_(sRef, vAddress, n_down_start, n_right_start))
    If Err.Number > 0 Then
        Debug.Print "Not Found"
        Exit Function
    End If
    
    If n_down_end = 0 Then n_down_end = xls_look_(rCell.Value, vLength)
    If n_right_end = 0 Then n_right_end = xls_look_(rCell.Value, vWidth)
    
    If n_down_end = 1 And n_right_end = 1 Then
        ReDim a_temp(1 To 1, 1 To 1) As Variant
        a_temp(1, 1) = rCell.Resize(n_down_end, n_right_end)
        a_make = a_temp
        Exit Function
    End If
    
    a_make = rCell.Resize(n_down_end, n_right_end)
End Function

'###################################################################
'   16.12.19 run fast
'   Turn off Screen Updating to speed up methods execution
'   Boolean input: TRUE function triggered ON, FALSE function triggered OFF
'-------------------------------------------------------------------
Sub xls_fast_(Optional status As Boolean = True)
    
    If status = True Then
        Application.Calculation = 0
        Application.ScreenUpdating = 0
    Else
        Cells(1, 1).Select          'is this necessary?
        Application.StatusBar = ""  'is this necessary?
        Application.Calculation = 1
        Application.ScreenUpdating = 1
    End If
End Sub

'###################################################################
'   16.12.19 : turns input box text into an array
'   08.11.21 : updated range input w/ ascending & descending
'              fixed limit of 1 & 2 digits in range input
'              added input box description
'-------------------------------------------------------------------
Function xls_inputbox_() As Variant

    Dim input_text As String
    Dim a_temp As Variant
    Dim range_start As Integer, range_end As Integer, i As Integer
    input_text = InputBox( _
                            "Type text you want to convert." & vbNewLine & _
                            "Single entries: separate w/ a comma." & vbNewLine & _
                            "Range: input boundaries separated with a dash.", _
                            "Turn input text into an array" _
                            )
        
    If InStr(input_text, "-") <> 0 Then
        a_temp = Split(input_text, "-")
        input_text = ""
        range_start = a_temp(0)
        range_end = a_temp(1)
        
        For i = 0 To Abs(range_end - range_start)
            If range_start < range_end Then
                input_text = input_text & range_start + i & ","
            ElseIf range_start > range_end Then
                input_text = input_text & range_start - i & ","
            Else
                xls_inputbox_ = range_start
                Exit Function
            End If
        Next

        xls_inputbox_ = Mid(input_text, 1, Len(input_text) - 1)
    End If

    xls_inputbox_ = Split(input_text, ",")
End Function
'###################################################################
'   24.09.20 a_base0: makes 1st row 0 base
'-------------------------------------------------------------------
Function a_base0(a_temp As Variant)

    Dim i As Integer, j As Integer
        
    ReDim a_temp1(0 To UBound(a_temp) - LBound(a_temp), LBound(a_temp, 2) To UBound(a_temp, 2)) As Variant
    
    For i = LBound(a_temp) To UBound(a_temp)
        For j = LBound(a_temp, 2) To UBound(a_temp, 2)
            a_temp1(i - LBound(a_temp), j) = a_temp(i, j)
        Next
    Next
    
    a_base0 = a_temp1

End Function
'###################################################################
'   26.01.15 a_base0: Creates a new worksheet
'            and pastes passed array
'-------------------------------------------------------------------
Sub a_test(aTemp As Variant)

    Dim sTxt As Variant
    
    'sTxt = UBound(aTemp)

    On Error Resume Next
        Sheets("test").Activate
        If Err.Number > 0 Then Sheets.Add.Name = "test"
    On Error GoTo 0
    
    Cells.Clear

    Call a_paste("B2", 0, 0, aTemp)
   
End Sub


