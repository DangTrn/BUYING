Attribute VB_Name = "ORDER_FORM"
Sub NSP()
    Dim FA23 As Workbook, book As Workbook
    Dim NSP As ListObject
    Dim store As Variant
    Dim path As String, name As String
    Dim i As Integer
    
    Set FA23 = Workbooks("FA23 BUYING - ADOPTION LIST.xlsm")
    Set NSP = FA23.Worksheets("FNL_ORDER").ListObjects("FINAL_ORDER")


    
    path = "D:\OneDrive\ACFC\NIKE - Documents\From G Suite Drive\5. MERCHANDISING TEAM\BUY-IN\FY24\FA23\5-ORDER FORM\FINAL ORDER\"
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    If Dir(path) <> "" Then
        Kill path & "*"
    End If
    
    store = FA23.Worksheets("DOOR PROFILE").ListObjects("DOOR_PROFILE").DataBodyRange.Value
    
    For i = LBound(store) To UBound(store)
        NSP.AutoFilter.ShowAllData
        NSP.Range.AutoFilter _
            Field:=1, _
            Criteria1:=store(i, 1)
        
        Workbooks.add.SaveAs Filename:=path & store(i, 5), FileFormat:=xlCSV
        Set book = Workbooks(store(i, 5) & ".csv")
        
        FA23.Worksheets("FNL_ORDER").Range("FINAL_ORDER[[#All],[Product Name]:[UPC]]").SpecialCells(xlCellTypeVisible).COPY
        
        With book.Worksheets(1)
            .Range("A1").PasteSpecial Paste:=xlPasteValues
            .Range("A1").PasteSpecial Paste:=xlPasteFormats
            .Range("A:AD").EntireColumn.AutoFit
        End With
        
        book.Close savechanges:=True
    Next i
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    
    MsgBox "DONE!"
    
End Sub
