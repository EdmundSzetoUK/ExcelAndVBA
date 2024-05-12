Attribute VB_Name = "CopyPasteRawDataModule"
Sub CopyPasteRawData()
    Dim lRow As Integer
    Dim strWbName, strWsName As String
    Dim strTodayDate, strYesterdayDate As String
    Dim oFso As Object
    Dim strFileName As String
    Dim strFilePath As String
    Dim strFolderPath As String
    Dim strFileExists As String

    ' Get the current date in the format "yyyymmdd"
    strTodayDate = Format(Date, "yyyymmdd")
    strYesterdayDate = Format(Date - 1, "yyyymmdd")

    ' Current Workbook
    strReportName = ActiveWorkbook.Name
    strWsName = ActiveWorkbook.ActiveSheet.Name
    
    ' -----------------------------------------
    ' Copy negative raw data
    
    'Check File exist
    strFileName = "Negative Report " & strTodayDate & ".xls"
    strFilePath = Environ("USERPROFILE") & "\Downloads\" & strFileName
    strFileExists = Dir(strFilePath)
    
    If strFileExists = "" Then
        ' If the filename is not match, then use open file dialog
        ' Construct the folder path
        strFolderPath = Environ("USERPROFILE") & "\Downloads\"
        With Application.FileDialog(msoFileDialogFilePicker)
            .InitialFileName = strFolderPath
            .Title = "Select a File to Attach"
            .Filters.Add "All excel", "*.xlsx", 1
            .AllowMultiSelect = False
            .InitialFileName = "The Report*.*"
            If .Show = -1 Then
                strFilePath = .SelectedItems(1)
            Else
                Exit Sub
            End If
        End With

        ' Get raw data filename
        Set oFso = CreateObject("Scripting.FileSystemObject")
        strFileName = oFso.GetFileName(strFilePath)
    
    End If
    
    ' Open negative raw data file
    Workbooks.Open strFilePath
    
    ' Goto Raw Data WorkBook
    Workbooks(strFileName).Worksheets("Sheet1").Activate
    lRow = Range("A3").End(xlDown).Row
    
    ' Add Filter
    Selection.AutoFilter
    ' Filter by the ID < 100
    ActiveSheet.Range("$A$1:$M$" & lRow).AutoFilter Field:=1, Criteria1:="<100", Operator:=xlAnd

    ' Select all Filtered Data and Copy
    lRow = Range("A3").End(xlDown).Row
    Range("A3:M" & lRow).Select
    Selection.Copy
    
    '---------------------------------------------
    ' Paste the raw data to the report
    
    ' Goto Negative Report
    Workbooks(strReportName).Worksheets(strWsName).Activate
    
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
              
    '---------------------------------------------
    ' Clean up the objects
    Set oFso = Nothing
    ' Close raw data file
    Application.DisplayAlerts = False
    Workbooks(strFileName).Close SaveChanges:=False
    Application.DisplayAlerts = True
 
End Sub



