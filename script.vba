Sub SendEmailsFromTextE2()
    Dim OutlookApp As Object, Mail As Object
    Dim wsC As Worksheet, wsT As Worksheet
    Dim i As Long, lastRow As Long
    Dim email As String, subjectKey As String
    Dim sSubject As String, sBody As String
    Dim msgRow As Range

    Set wsC = ThisWorkbook.Sheets("Contacts")
    Set wsT = ThisWorkbook.Sheets("Text")
    
    subjectKey = wsT.Range("E2").Value
    If subjectKey = "" Then
        MsgBox "Subject key in E2 is empty.", vbExclamation
        Exit Sub
    End If

    ' Find the row matching the subject key in "Text"
    Set msgRow = wsT.Range("A2:A" & wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row) _
                     .Find(What:=subjectKey, LookIn:=xlValues, LookAt:=xlWhole)
    
    If msgRow Is Nothing Then
        MsgBox "Subject '" & subjectKey & "' not found in 'Text' sheet.", vbExclamation
        Exit Sub
    End If

    sSubject = msgRow.Offset(0, 1).Value ' sMessage (column B)
    sBody = msgRow.Offset(0, 2).Value    ' message (column C)

    Set OutlookApp = CreateObject("Outlook.Application")
    lastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        email = wsC.Cells(i, 1).Value
        If email <> "" Then
            Set Mail = OutlookApp.CreateItem(0)
            With Mail
                .To = email
                .Subject = sSubject
                .Body = sBody
                .Display ' Change to .Send to send directly
            End With
        End If
    Next i
End Sub
