VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnImport_Click()
 Dim strNameSheet As String
 Dim strPathFile As String
 Dim lngCount As Long
 Dim lngIndexDot As Long

 With Application.FileDialog(msoFileDialogOpen)
 .Title = "Select File"
 .AllowMultiSelect = False
 .Show
 strPathFile = .SelectedItems(1)
 End With

 strNameSheet = Dir(strPathFile, vbNormal)
 lngIndexDot = InStr(strNameSheet, ".")
 strNameSheet = Left(strNameSheet, lngIndexDot - 1)
 Worksheets.Add Count:=1, After:=Worksheets(1)
 ActiveSheet.Name = strNameSheet
 With Worksheets(strNameSheet).QueryTables.Add(Connection:="TEXT;" & _
 strPathFile, Destination:=Worksheets(strNameSheet).Range("A1"))
 .Name = "Time Attendant"
 .FieldNames = True
 .RowNumbers = False
 .FillAdjacentFormulas = False
 .PreserveFormatting = True
 .RefreshOnFileOpen = False
 .RefreshStyle = xlInsertDeleteCells
 .SavePassword = False
 .SaveData = True
 .AdjustColumnWidth = True
 .RefreshPeriod = 0
 .TextFilePromptOnRefresh = False
 .TextFilePlatform = 874
 .TextFileStartRow = 1
 .TextFileParseType = xlDelimited
 .TextFileTextQualifier = xlTextQualifierDoubleQuote
 .TextFileConsecutiveDelimiter = False
 .TextFileTabDelimiter = True
 .TextFileSemicolonDelimiter = False
 .TextFileCommaDelimiter = False
 .TextFileSpaceDelimiter = False
 .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
 .TextFileTrailingMinusNumbers = True
 .Refresh BackgroundQuery:=False
 End With
End Sub

Private Sub btnPivotTable_Click()
    Const pivotTable_sheet = "PivotTable"
    Worksheets(pivotTable_sheet).Activate
    ActiveSheet.Range("A:Z").Value = Empty
    
    'ActiveSheet.Range("A1").Value = "Sum of จำนวนเงินที่จ่าย"
End Sub
