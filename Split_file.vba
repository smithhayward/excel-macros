Sub Split_File()
' AUTHOR: SMITH HAYWARD
' DATE: 1/14/2025
' TESTING: Executed with a file containing over 10,000 records with a records per file of 999 (1,000 including headers) which created 11 files successfully.
  
' HARD CODED TO TAKE A 3-COLUMN WORKSHEET AND CREATE SEPARATE WORKBOOKS WITH 1000 records (Header inclusive) each.
'
' INSTRUCTIONS:  Add this VBA Macro into your macro-enabled (.xlsm) excel file and run the Split_File macro.  If your needs are different you can modify the header collection and writing section as well as change the number of records per file.
'
' ==============================================
'   VARIABLE DECLARATION AND INITIALIZATION
' ==============================================
Dim filePath As String
filePath = ThisWorkbook.Path
Dim filePrefix As String
filePrefix = InputBox("Give the series of files a unique file name prefix:")

    If filePrefix = "" Then
    filePrefix = "splitfile"
    End If

Dim fileCount As Integer
fileCount = 1

Dim fileExt As String
fileExt = "xlsx"

Dim savePath As String

Dim header(2) As String
Dim data(99, 999, 2) As String

' ==============================================
' CAPTURE HEADER VALUES
' ==============================================

With ThisWorkbook.Sheets(1)
    
    header(0) = .Cells(1, 1).Value
    header(1) = .Cells(1, 2).Value
    header(2) = .Cells(1, 3).Value

End With

' ==============================================
' CAPTURE DATA PAYLOAD
' ==============================================
Dim currentRow As Integer
currentRow = 2
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim lastFileIndex As Integer
Dim dataEnd As Integer
dataEnd = 0
Dim recordsPerFile As Integer
recordsPerFile = 999

For z = 0 To 100

    For y = 1 To recordsPerFile
    
        For x = 0 To 2
            data(z, y, x) = ThisWorkbook.Sheets(1).Cells(currentRow, x + 1).Value
            
        Next x
        
        currentRow = currentRow + 1
        If ThisWorkbook.Sheets(1).Cells(currentRow, 1).Value = "" And ThisWorkbook.Sheets(1).Cells(currentRow, 2).Value = "" And ThisWorkbook.Sheets(1).Cells(currentRow, 3).Value = "" Then
            dataEnd = 1
            Exit For
        End If
        
    
    Next y

    If dataEnd = 1 Then
        lastFileIndex = z
        Exit For
    End If
    
Next z




' ==============================================
' PRIMARY FILE GEN LOOP
' ==============================================

Dim NewWb As Workbook

For z = 0 To lastFileIndex

    Set NewWb = Workbooks.Add
    savePath = filePath + "\" + filePrefix + "_" + CStr(z + 1) + "." + fileExt
    
    ' ADD THE HEADER
        NewWb.Sheets(1).Cells(1, 1).Value = header(0)
        NewWb.Sheets(1).Cells(1, 2).Value = header(1)
        NewWb.Sheets(1).Cells(1, 3).Value = header(2)
    
    
    ' WRITE THE DATA
    For y = 1 To recordsPerFile
        
            For x = 0 To 2
                NewWb.Sheets(1).Cells(y + 1, x + 1).Value = data(z, y, x)
                    
            Next x
        
    Next y
    
    
    On Error Resume Next
        NewWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "Error saving the file. Please check the save path.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    NewWb.Close SaveChanges:=False


Next z

' ALL FILES CREATED AND CLOSED | COMPLETE
Dim plural As String
plural = ""
haveHas = "has"
If lastFileIndex > 0 Then
plural = "s"
haveHas = "have"
End If

MsgBox CStr(lastFileIndex + 1) + " file" + plural + " " + haveHas + " been generated."


End Sub
