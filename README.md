<img align="center" src="/image/VBA-Banner.jpg" alt="Banner" />

# Excel And Excel VBA
This knowledge of Excel and Excel VBA sharing from my working experience. 

## A. Download Email Attachment
### Objective
The DownloadEmailAttachment function to scraping and cleaning data. 
* It sets up Outlook variables and connects to the MAPI namespace.
* Navigating to the designated folder, it checks each email (newest to oldest) for specified criteria. If an email meets the criteria (received today, subject: "NEGATIVE REPORT"), it saves the attachment "Negative Report.xls" to a folder. A message box then confirms the outcome.

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

## B. Copy Raw Data
### Objective
Automates the extraction and processing of data from a raw data file. 
* It first checks for the file's existence; if not found, it prompts user input. After identifying the file, it filters the data based on a specified condition, selecting only relevant information. 
* The filtered data is then copied and pasted into a report, streamlining the analysis and reporting process. Cleanup operations ensure proper resource management, closing temporary files after processing.
* This automation enhances efficiency and accuracy in data management, facilitating informed decision-making by seamlessly integrating disparate data sources into actionable insights.

The Excel Marco Code: [CopyPasteRawData](/CopyPasteRawData.bas)

### Detail
* Copying Negative Raw Data:
  * This part of the code deals with extracting specific data from a file.
It first checks if a file with a certain name exists. If it doesn't, it prompts the user to select the file manually.
Once the file is identified, it opens it and navigates to a specific worksheet within that file.
  * It then applies a filter to select only the relevant data, in this case, data where a certain condition is met (in this case, where the value in the first column is less than 100).
  * After filtering, it selects and copies the filtered data.
* Pasting Raw Data to the Negative Report:
  * This part of the code deals with pasting the extracted data into another location, specifically a report.
It raw data to the report file and worksheet where the data needs to be pasted.
  * It then pastes the copied data into a specific cell range, ensuring that only the values are pasted without any formatting.
* Cleanup:
  * This part of the code ensures that any resources or objects used during the process are properly closed or released.
  * It cleans up by closing the file containing the raw data and resetting any temporary variables used during the process.
