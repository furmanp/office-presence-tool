Attribute VB_Name = "Logs_Module"
'@Folder("Monthly_Statement")
Option Explicit
Sub DownloadMonthlyLog()

    Dim monthNumber As Long
    Dim oCalendar As Outlook.Folder
    Dim oItems As Outlook.Items
    Dim oResItems As Outlook.Items
    Dim oPresItems As Outlook.Items
    Dim oAppt As Outlook.AppointmentItem
    Dim strRestriction As String
    Dim myStart As Date
    Dim myEnd As Date
    
    Application.ScreenUpdating = False
    
    Worksheets("Logs").Range("A4:B34").Clear

    monthNumber = month(DateValue("01-" & Worksheets("Logs").Cells(2, 3) & "-1900"))    'convert month name from cell to its number
    myStart = DateSerial(Year(Date), monthNumber, 1)                                    'create first day of the month in a Date format
    myEnd = DateAdd("m", 1, myStart)                                                    'calculates the first day of the next month
    myEnd = DateAdd("d", -1, myEnd)                                                     'substracts one day from previously calculated Date (resulting in last day of the month)
    
    Set oCalendar = Outlook.Application.Session.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.Items
    oItems.Sort "[Start]"

    strRestriction = "[Start] <= '" & Format$(myEnd, "dd/mm/yyyy hh:mm AMPM") _
      & "' AND [End] >= '" & Format(myStart, "dd/mm/yyyy hh:mm AMPM") & "'"
         
    Set oResItems = oItems.Restrict(strRestriction)                                     'Restrict the Items collection
    Set oPresItems = oResItems.Restrict("[Categories] = 'OfficePresence' ")             'Restrict the Items Category
    
    Dim i&, j&
     
    For i = 0 To Day(myEnd) - 1
        Worksheets("Logs").Cells(4 + i, 1) = myStart + i
        
        If weekDay(myStart + i) = 1 Or weekDay(myStart + i) = 7 Then                    'color fill cells that show Weekends
            Worksheets("Logs").Cells(4 + i, 1).Interior.ColorIndex = 40
        End If
        
        For Each oAppt In oPresItems
            With oAppt
                If .Start = myStart + i Then
                    For j = 0 To (.Duration / 1440) - 1
                        Worksheets("Logs").Cells(4 + i + j, 2) = .Subject
                        If .BusyStatus = olOutOfOffice Then
                            Worksheets("Logs").Cells(4 + i, 2).Interior.Color = RGB(235, 80, 50)
                            Worksheets("Logs").Cells(4 + i, 2).Font.Italic = True
                        End If
                    Next
                End If
            End With
        Next
    Next
    
    Set oCalendar = Nothing
    Set oItems = Nothing
    Set oResItems = Nothing
    Set oPresItems = Nothing
    Set oAppt = Nothing
    
    Application.ScreenUpdating = True
    
End Sub

