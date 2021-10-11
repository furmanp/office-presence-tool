Attribute VB_Name = "Calendar_Module"
'@Folder("Appointment_Creator")
Option Explicit

Sub SetOfficePresence(WeekNumber As Long, yearNum As Long)
    
    Dim myStart As Date
    Dim myEnd As Date
    Dim oOutlook As Outlook.Application
    Dim oNamespace As Outlook.Namespace
    Dim oCalendar As Outlook.Folder
    Dim oItems As Outlook.Items
    Dim oAppt As Outlook.AppointmentItem
    Dim strRestriction As String
        
    Dim i As Long
    
    CreateCustomCategory ("OfficePresence")
        
    myStart = GetDayFromWeekNumber(yearNum, WeekNumber)
    myEnd = DateAdd("d", 4, myStart)
    
    Set oOutlook = Outlook.Application
    Set oNamespace = oOutlook.GetNamespace("MAPI")
    Set oCalendar = oNamespace.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.Items
     
    oItems.IncludeRecurrences = False
    
     
    strRestriction = "[Start] <= '" & Format$(myEnd, "dd/mm/yyyy hh:mm AMPM") _
      & "' AND [End] >= '" & Format(myStart & " 0:01am", "dd/mm/yyyy hh:mm AMPM") & "'"
         
    
    Set oItems = oItems.Restrict(strRestriction)                                     'Restrict the Items collection
    Set oItems = oItems.Restrict("[Categories] = 'OfficePresence' ")                 'Restrict the Items Category
    
    oItems.Sort "[Start]"
    
Line1:
    If oItems.Count = 0 Then
        For i = 0 To 4
            Call CreateAppt(myStart + i, i, "OfficePresence")
        Next i
    Else
        For i = oItems.Count To 1 Step -1
            Set oAppt = oItems.Item(i)
                oAppt.Delete
        Next i

        GoTo Line1
    End If
    
    Set oCalendar = Nothing
    Set oItems = Nothing
    Set oAppt = Nothing
End Sub

Private Sub CreateAppt(oDate As Date, weekDay As Long, oCategoryName As String)

    Dim oAppt As AppointmentItem
                  
    Set oAppt = Outlook.Application.CreateItem(olAppointmentItem)
   
    If IsEmpty(Cells(weekDay + 3, 3)) = False Then
        With oAppt
            .Subject = CStr(Cells(weekDay + 3, 3))
            If CheckBusyStatus(.Subject) Then
                .BusyStatus = olOutOfOffice
            End If
            .Start = oDate
            .AllDayEvent = True
            .ReminderSet = False
            .Categories = oCategoryName
            .Save
        End With
    End If
End Sub

Private Sub CreateCustomCategory(oCategoryName As String)
    Dim objNameSpace As Namespace
    Dim objCategory As Category
    
    Set objNameSpace = Outlook.Application.GetNamespace("MAPI")                     ' Obtain a NameSpace object reference.
                                                                                    ' Check if the Categories collection for the Namespace
    For Each objCategory In objNameSpace.Categories                                 ' contains one or more Category objects.
        If objCategory.Name = oCategoryName Then                                    'If category exists, update its color only and exit Sub
            objCategory.Color = CategoryColor(Worksheets("Setup").Cells(6, 3))
            Exit Sub
        End If
    Next
    
    Set objCategory = objNameSpace.Categories.Add(oCategoryName)                    'If category doesn't exist, create it and give it a color
    objCategory.Color = CategoryColor(Worksheets("Setup").Cells(6, 3))
    
    Set objCategory = Nothing                                                       ' Clean up.
    Set objNameSpace = Nothing
End Sub

Private Function GetDayFromWeekNumber(InYear As Long, WeekNumber As Long, Optional DayInWeek1Monday7Sunday As Long = 1) As Date
    Dim i As Long: i = 1

    Do While weekDay(DateSerial(InYear, 1, i), vbMonday) <> DayInWeek1Monday7Sunday
        i = i + 1
    Loop

    GetDayFromWeekNumber = DateAdd("ww", WeekNumber - 1, DateSerial(InYear, 1, i))
End Function

Private Function CheckBusyStatus(statusKeyword As String) As Boolean
    
    CheckBusyStatus = False
    
    Dim keywordArray(1 To 5) As Variant
    keywordArray(1) = "OFF"
    keywordArray(2) = "ILL"
    keywordArray(3) = "SICK"
    keywordArray(4) = "VACATION"
    keywordArray(5) = "HOLIDAYS"
    
    Dim i%
    
    For i = 1 To 5
        If UCase(statusKeyword) = keywordArray(i) Then
            CheckBusyStatus = True
            Exit Function
        End If
    Next
End Function

Private Function CategoryColor(cColor As String)
    Dim colorCode As Long
    
    Select Case cColor
    Case "Blue"
        colorCode = 8
    Case "Navy Blue"
        colorCode = 23
    Case "Dark Gray"
        colorCode = 14
    Case "Dark Green"
        colorCode = 20
    Case "Dark Pink"
        colorCode = 25
    Case "Lime Green"
        colorCode = 22
    Case "Dark Orange"
        colorCode = 17
    Case "Brown"
        colorCode = 18
    Case "Dark Purple"
        colorCode = 24
    Case "Dark Red"
        colorCode = 16
    Case "Gold"
        colorCode = 19
    Case "Gray"
        colorCode = 13
    Case "Green"
        colorCode = 5
    Case "Orange"
        colorCode = 2
    Case "Peach"
        colorCode = 3
    Case "Lavender"
        colorCode = 9
    Case "Red"
        colorCode = 1
    Case "Light Gray"
        colorCode = 11
    Case "Teal"
        colorCode = 6
    Case "Yellow"
        colorCode = 4
    End Select

    CategoryColor = colorCode
End Function


