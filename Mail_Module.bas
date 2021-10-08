Attribute VB_Name = "Mail_Module"
'@Folder("KOW_Mail")
Option Explicit
Sub Mail_KOW(weekNum As Long)

    Dim OutApp As Object
    Dim OutMail As Object
    Dim mailTitle As String
    Dim officePresence(1 To 5) As String
    Dim i&
 
    For i = 1 To 5                                                                                  'Convert users input into 5 (workdays) array
        officePresence(i) = CStr(Cells(i + 2, 2)) + " " + CStr(Cells(i + 2, 3))
    Next i
    
    mailTitle = "KOW " + CStr(weekNum)

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = Worksheets("Setup").Cells(2, 3)
        .CC = ""
        .BCC = ""
        .Subject = mailTitle
        .Body = "Hello everyone," & vbNewLine & vbNewLine & "My week goes as follows:" & vbNewLine & _
                officePresence(1) & vbNewLine & _
                officePresence(2) & vbNewLine & _
                officePresence(3) & vbNewLine & _
                officePresence(4) & vbNewLine & _
                officePresence(5) & vbNewLine & vbNewLine & _
                CStr(Worksheets("Setup").Cells(8, 3)) & vbNewLine & _
                CStr(Worksheets("Setup").Cells(4, 3))
        .Send
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub


