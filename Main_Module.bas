Attribute VB_Name = "Main_Module"
'@Folder("KOW_Sender")

Option Explicit

Sub Main()
    
    Dim weekNum&, yearNum&
    
    Worksheets("Sender").Activate
'=======================================================================================================================================
    If IsEmpty(Cells(2, 3)) Then                                                'check if Calendar week is set for Default or Custom
        weekNum = WorksheetFunction.weekNum(Date, 21)                           'If Empty then CW is Default (taken from Today's date)
    Else
        weekNum = CInt(Cells(2, 3))                                             'If not empty, use user defined CW
    End If
    
    If IsEmpty(Cells(2, 6)) Then                                                'check if Year is set for Default of Custom
        yearNum = Year(Date)
    Else
        yearNum = CInt(Cells(2, 6))
    End If
'=======================================================================================================================================
    Select Case Worksheets("Sender").Cells(8, 3)                                'Execute set of methods based on user choice
        Case "KOW + Calendar"                                                   'KOW Only / KOW + Calendar / Calendar Only
            Call Mail_KOW(weekNum)
            Call SetOfficePresence(weekNum, yearNum)
            
        Case "KOW Only"
            Call Mail_KOW(weekNum)
            
        Case "Calendar Only"
            Call SetOfficePresence(weekNum, yearNum)
        
    End Select
End Sub



