Attribute VB_Name = "Module1"
Sub RearrangeColumnsByHeaders()
    Dim ws As Worksheet
    Dim sourceColumns As Range
    Dim targetColumns As Range
    Dim headerList As Variant
    Dim header As Variant
    Dim i As Long
    
    ' Set the worksheet where the columns are located
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    
    ' Define the list of headers in the desired order
    headerList = Array(1, 1, 1, 1) ' Adjust as needed
    
    ' Set the source columns based on the header list
    Set sourceColumns = ws.Range("A1").Resize(, UBound(headerList) + 1).EntireColumn
    
    ' Set the target columns based on the header list
    Set targetColumns = ws.Range("A1").Resize(, UBound(headerList) + 1)
    
    ' Rearrange columns based on the header list
    For i = LBound(headerList) To (UBound(headerList))
        header = 2 ' (UBound(headerList) + 1) - i
        If header <> 1 Then
         sourceColumns.Columns(header).Cut
         sourceColumns.Columns(1).Insert Shift:=xlToLeft ' once put in front it is no longer included in source columns
        ' Application.CutCopyMode = False
        Application.Wait Now + TimeValue("00:00:05")
        End If
    Next i
End Sub

