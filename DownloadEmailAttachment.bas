Attribute VB_Name = "DownloadEmailAttachmentModule"
' Enum for download email attachment result
Private Enum eDlEmailAttach
 Success = 0
 NoNewEmail = -1
 EmailSubjectNotMatch = -2
 EmailFolderNotMatch = -3
End Enum

Sub DownloadEmailAttachment()
    Dim olApp As Object
    Dim OutlookNamespace As Object
    Dim OutAccount As Object
    Dim EmailAcctNo As Integer
    
    Dim Folder As Object
    Dim MailItem As Object
    Dim Attachment As Object
    Dim SaveFolder As String
    Dim fileName As String
    Dim TodayDate As String
    Dim iResult As eDlEmailAttach
    
    ' Define the save folder using the user's profile folder
    SaveFolder = Environ("USERPROFILE") & "\Downloads\" ' Save in the user's profile "Downloads" folder
    ' Get the current date in the format "yyyymmdd"
    TodayDate = Format(Date, "yyyymmdd")
    ' Create an instance of Outlook
    Set olApp = CreateObject("Outlook.Application")
    ' Get the MAPI namespace
    Set OutlookNamespace = olApp.GetNamespace("MAPI")
    
    ' Open the edmund.s@jegroupltd.co.uk Mail Box -> Folders Inbox > Reporting > QlikView
    Set Folder = OutlookNamespace.Folders("xxx.xx@xxxltd.co.uk").Folders("Inbox").Folders("Reporting").Folders("QlikView")
    iResult = EmailFolderNotMatch
    ' Loop through each mail item in reverse order
    For i = Folder.Items.Count To 1 Step -1
        Set MailItem = Folder.Items(i)
        ' Check if the email was received today
        If DateDiff("d", MailItem.CreationTime, Now) = 0 Then
            ' Check the email subject and attachment name 'NEGATIVE REPORT'
            If InStr(MailItem.Subject, "NEGATIVE REPORT") <> 0 Then
                ' Find all attachment
                For Each Attachment In MailItem.Attachments
                    ' Check attachment name
                    If Attachment.DisplayName = "Negative Report.xls" Then
                        ' Save the attachment with a new file name
                        fileName = "Negative Report " & TodayDate & ".xls"
                        Attachment.SaveAsFile SaveFolder & fileName
                        iResult = Success
                        Exit For ' Exit the loop once the attachment is saved
                    End If
                Next Attachment
                
                ' Exit the loop if an older email is encountered
                Exit For
            Else
                iResult = EmailSubjectNotMatch
            End If
        Else
            ' Exit the loop if an older email is encountered
            iResult = NoNewEmail
            Exit For
        End If
    Next i
     
    ' Popup Message
    If iResult = Success Then
        MsgBox "Successful to download attachment to " & SaveFolder & fileName, vbOKOnly
    ElseIf iResult = EmailSubjectNotMatch Then
        MsgBox "Negative Report email is not found.", vbOKOnly, "Warning"
    ElseIf iResult = NoNewEmail Then
        MsgBox "No any new emails today!", vbOKOnly, "Warning"
    ElseIf iResult = EmailFolderNotMatch Then
        ' Display a message if the attachment was not found
        MsgBox "Email account or folders is not match", vbOKOnly, "Warning"
    End If
    
    ' Clean up the objects
    Set Attachment = Nothing
    Set MailItem = Nothing
    Set Folder = Nothing
    Set OutlookNamespace = Nothing
    Set olApp = Nothing
End Sub