<img align="center" src="/image/VBA-Banner.jpg" alt="Banner" />

# Excel And Excel VBA
This knowledge of Excel and Excel VBA sharing from my working experience. 

## Download Email Attachment
### Objective
The DownloadEmailAttachment function to scraping and cleaning data. 

It sets up Outlook variables and connects to the MAPI namespace. Navigating to the designated folder, it checks each email (newest to oldest) for specified criteria. If an email meets the criteria (received today, subject: "NEGATIVE REPORT"), it saves the attachment "Negative Report.xls" to a folder. A message box then confirms the outcome.

The Excel Marco Code: [DownloadEmailAttachment](/DownloadEmailAttachment.bas)

### Detail
* Initialize necessary variables:
  * Declare variables for Outlook application (olApp), namespace (OutlookNamespace), email account (OutAccount), and other relevant objects.
Define variables for the email folder (Folder), individual email item (MailItem), attachment (Attachment), save folder path (SaveFolder), file name (fileName), today's date (TodayDate), and the result of the download operation (iResult).
* Define the save folder path:
  * Set the SaveFolder variable to the user's profile "Downloads" folder using the Environ("USERPROFILE") function.
* Get the current date:
  * Retrieve today's date in the format "yyyymmdd" using the Format(Date, "yyyymmdd") function and assign it to the TodayDate variable.
* Create an instance of Outlook:
  * Use the CreateObject("Outlook.Application") function to create an instance of the Outlook application and assign it to the olApp variable.
* Get the MAPI namespace:
  * Use the GetNamespace("MAPI") method of the olApp object to retrieve the MAPI namespace and assign it to the OutlookNamespace variable.
* Set the target email folder:
  * Navigate to the specified folder (xxx.xx@xxx.co.uk\Inbox\Reporting\QlikView) within the mailbox using the Folders property of the OutlookNamespace object and assign it to the Folder variable.
* Iterate through each email:
  * Loop through each email item in reverse order (from newest to oldest) within the specified folder using a For loop with a step of -1.
* Check email criteria:
  * For each email item, check if it was received today (DateDiff("d", MailItem.CreationTime, Now) = 0) and if its subject contains "NEGATIVE REPORT" (InStr(MailItem.Subject, "NEGATIVE REPORT") <> 0).
* Download attachment:
  * If the email meets the criteria, iterate through its attachments and check if the attachment's display name matches "Negative Report.xls". If found, save the attachment with a new file name (fileName) in the specified save folder (SaveFolder).
