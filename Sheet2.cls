VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const receipt_sheet = "Receipt"
Const receipt_db_sheet = "DataImported"
Private Sub search_receipt_btn_Click()

    Dim key As String
    key = receipt_no_txtbox.Text
    Dim checkName As Boolean
    checkName = False
    
    Dim nameCol As Integer
    nameCol = 2
    Dim val As String
    
        
    Do While Worksheets("DataImported").Range("B" & nameCol).Value <> ""
        val = Worksheets("DataImported").Range("B" & nameCol).Value
        
        If InStr(val, key) <> 0 Then
            checkName = True
            Exit Do
        End If
        nameCol = nameCol + 1
    Loop
        
    If checkName Then
     Call search_receipt(nameCol)
    Else
        Call clearReceipt
        MsgBox "Customer name must be a-y only !!", vbOKOnly + vbCritical, "Be careful"
    End If

End Sub


Public Sub search_receipt(numRow)
    Dim key As String
    Dim val As String
    Dim rRow As Integer
    Dim nameCol As Integer
    
    nameCol = 2
    rRow = 16
    key = receipt_no_txtbox.Text
    
    Worksheets(receipt_sheet).Range("F9").Value = Worksheets(receipt_db_sheet).Range("A" & numRow).Value
    Worksheets(receipt_sheet).Range("F10").Value = Worksheets(receipt_db_sheet).Range("H" & numRow).Value
    Worksheets(receipt_sheet).Range("C12").Value = Worksheets(receipt_db_sheet).Range("B" & numRow).Value
    Worksheets(receipt_sheet).Range("C13").Value = Worksheets(receipt_db_sheet).Range("C" & numRow).Value
    
     Do While Worksheets("DataImported").Range("B" & nameCol).Value <> ""
        val = Worksheets("DataImported").Range("B" & nameCol).Value
        If InStr(val, key) <> 0 Then
            Worksheets(receipt_sheet).Range("C" & rRow).Value = Worksheets(receipt_db_sheet).Range("D" & nameCol).Value
            Worksheets(receipt_sheet).Range("D" & rRow).Value = Worksheets(receipt_db_sheet).Range("F" & nameCol).Value
            Worksheets(receipt_sheet).Range("E" & rRow).Value = Worksheets(receipt_db_sheet).Range("E" & nameCol).Value
            Worksheets(receipt_sheet).Range("F" & rRow).Value = Worksheets(receipt_db_sheet).Range("F" & nameCol).Value * Worksheets(receipt_db_sheet).Range("E" & nameCol).Value
            
            rRow = rRow + 1
        End If
        nameCol = nameCol + 1
    Loop
End Sub

Public Sub clearReceipt()
    Worksheets(receipt_sheet).Range("F9").Value = ""
    Worksheets(receipt_sheet).Range("F10").Value = ""
    Worksheets(receipt_sheet).Range("C12").Value = ""
    Worksheets(receipt_sheet).Range("C13").Value = ""
    
    Dim i As Integer
    For i = 16 To 25
        Worksheets(receipt_sheet).Range("C" & i).Value = ""
        Worksheets(receipt_sheet).Range("D" & i).Value = ""
        Worksheets(receipt_sheet).Range("E" & i).Value = ""
        Worksheets(receipt_sheet).Range("F" & i).Value = ""
    Next i
End Sub
