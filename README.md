# 📧 Bulk Email Sender via Excel and Outlook (VBA)

This VBA script automates the process of sending individualized emails via Microsoft Outlook using data stored in an Excel workbook. It’s optimized for simplicity and practical business use — no fluff, just function.

---

## 🗂️ Sheet Setup

### `Contacts` Sheet

| A (Email)       | B (First Name) | C (Last Name) |
|----------------|----------------|----------------|
| example@xyz.com | John           | Doe            |
| ...             | ...            | ...            |

- Only **Column A** (Email) is required.
- First and Last Name columns are optional and unused by this version of the script.

---

### `Text` Sheet

| A (Subject Key) | B (sMessage - Email Subject) | C (Message - Email Body) | E (Trigger Cell) |
|----------------|------------------------------|---------------------------|------------------|
| Reminder       | Don't Forget!                | You have a meeting soon. | Reminder         |

- **E2** must contain the subject key you want to use.
- The script will find the first match of this key in Column A and use the corresponding `sMessage` and `message`.

---

## ⚙️ How to Use

1. Open your Excel workbook with the above sheet structure.
2. Press `Alt + F11` to open the **VBA Editor**.
3. Insert a new module (`Insert > Module`).
4. Paste the script below into the module:
   
   ```vba
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
                   .Display ' Change to .Send to send without preview
               End With
           End If
       Next i
   End Sub
