Sub ExportSharedInboxEmailsToCSV()
    ' === CONFIGURATION VARIABLES ===
    Dim sharedMailboxName As String: sharedMailboxName = "Your Shared Mailbox Name" ' e.g., "Blue Jeans Comapny"
    Dim folderName As String: folderName = "Inbox"
    Dim outputFileName As String: outputFileName = "shared-inbox.csv"
    Dim outputFolder As String: outputFolder = "Your Output Folder Path\" ' e.g., "C:\Users\YourName\Documents\"
    Dim bodyTruncateLength As Long: bodyTruncateLength = 1000

    ' === OUTLOOK OBJECTS ===
    Dim ns As Outlook.NameSpace
    Dim inboxFolder As Outlook.Folder
    Dim emailItem As Outlook.MailItem
    Dim emailItemCount As Long

    ' === FILE OBJECTS ===
    Dim csvFilePath As String
    Dim csvFile As Integer

    ' === EMAIL FIELDS ===
    Dim i As Long
    Dim emailSubject As String
    Dim emailSender As String
    Dim emailDate As String
    Dim emailBody As String

    ' === SETUP ===
    Set ns = Application.GetNamespace("MAPI")
    On Error Resume Next
    Set inboxFolder = ns.Folders(sharedMailboxName).Folders(folderName)
    On Error GoTo 0

    If inboxFolder Is Nothing Then
        MsgBox "Could not find the folder '" & folderName & "' under mailbox '" & sharedMailboxName & "'.", vbCritical, "Folder Not Found"
        Exit Sub
    End If

    csvFilePath = outputFolder & outputFileName
    csvFile = FreeFile
    Open csvFilePath For Output As csvFile

    ' === HEADER ROW ===
    Print #csvFile, "Subject,Sender,Date,Body"

    ' === LOOP THROUGH EMAILS ===
    emailItemCount = inboxFolder.Items.Count
    For i = 1 To emailItemCount
        On Error Resume Next
        Set emailItem = inboxFolder.Items(i)

        If Not emailItem Is Nothing Then
            If TypeOf emailItem Is MailItem Then
                emailSubject = Replace(emailItem.Subject, ",", " ")
                emailSender = Replace(emailItem.SenderName, ",", " ")
                emailDate = Replace(emailItem.ReceivedTime, ",", " ")
                emailBody = Replace(emailItem.Body, ",", " ")

                If Len(emailBody) > bodyTruncateLength Then
                    emailBody = Left(emailBody, bodyTruncateLength) & "..."
                End If

                Print #csvFile, """" & emailSubject & """,""" & emailSender & """,""" & emailDate & """,""" & emailBody & """"
            End If
        End If
    Next i

    Close csvFile
    MsgBox "Export complete! Emails saved to: " & csvFilePath, vbInformation
End Sub
