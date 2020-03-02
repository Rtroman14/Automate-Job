Sub systemLineItems()


    ' MASTER TEMPLATE variables
    Dim masterTemplate As Workbook
    Dim MASTER As Worksheet
    Set masterTemplate = Workbooks("MASTER TEMPLATE.xlsm")
    Set MASTER = masterTemplate.Worksheets("MASTER")


    ' RunQuery variables
    Dim runQuery As Workbook
    Dim runQuerySheet As Worksheet
    Set runQuery = Workbooks("RunQuery")
    Set runQuerySheet = runQuery.Worksheets("RunQuery")


    'System name for link
    Dim systemName As String
    Dim findDash As Long
    Dim rowNum


    'Count total number of systems in "RunQuery" Worksheet
    Dim totalSystems As Integer
    totalSystems = Application.CountA(runQuerySheet.Range("B2:B1000"))


    'Count total number of systems in "master" Worksheet
    Dim totalSystemsListMaster As Integer
    totalSystemsListMaster = Application.CountA(MASTER.Range("B4:B1000"))
    

    'Create new Worksheet labeled "Asset System Line Items"
    Dim assetSystemLineItems As Worksheet
    Set assetSystemLineItems = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    assetSystemLineItems.Name = "Asset System Line Items"


    'Creating column headings to "Asset System Line Items" Worksheet
    assetSystemLineItems.Range("A1").Value = "ID"
    assetSystemLineItems.Range("B1").Value = "assetSystemID"
    assetSystemLineItems.Range("C1").Value = "costSourceCode"
    assetSystemLineItems.Range("D1").Value = "classification"
    assetSystemLineItems.Range("E1").Value = "codeLabel"
    assetSystemLineItems.Range("F1").Value = "description"
    assetSystemLineItems.Range("G1").Value = "quantity"
    assetSystemLineItems.Range("H1").Value = "unit"
    assetSystemLineItems.Range("I1").Value = "opTotal"


    'Change all cells to "text" so line item code first character "0" isn't lost
    Cells.Select
    Selection.NumberFormat = "@"


    'Loop through RunQuery list to extract name before "-"
    For x = 2 To totalSystems
        runQuerySheet.Cells(x, 2).Value = systemNameLink(x)
    Next


    Workbooks("MASTER TEMPLATE.xlsm").Sheets("MASTER").Activate

    ' add line items to assetSystem's worksheet
    On Error Resume Next

    For runQueryRow = 2 To totalSystems

        For masterRow = 4 To totalSystemsListMaster

            If runQuerySheet.Cells(runQueryRow, 2).Value = MASTER.Cells(masterRow, 1) Then

                For lineItem = 1 To numLineItem(masterRow)

                    lineItemRow = emptyCellRow(2)

                    assetSystemLineItems.Cells(lineItemRow, 2).Value = runQuerySheet.Cells(runQueryRow, 1)
                    assetSystemLineItems.Range("C" & lineItemRow, "I" & lineItemRow).Value = masterLineItem(masterRow, lineItem)

                Next

            End If

        Next

    Next

   
    'Change all cells back to "General"
    assetSystemLineItems.Select
    Cells.Select
    Selection.NumberFormat = "General"

    
    'resize column width
    assetSystemLineItems.Columns("B:B").EntireColumn.AutoFit
    assetSystemLineItems.Columns("C:C").EntireColumn.AutoFit
    assetSystemLineItems.Columns("E:E").EntireColumn.AutoFit
    assetSystemLineItems.Columns("F:F").ColumnWidth = 22
    
    
    'Count total number of systems in "Asset System Line Items" Worksheet
    Dim assetSystemLineItemList
    assetSystemLineItemList = Application.CountA(assetSystemLineItems.Range("B1:B1000"))


    ' highlight systems copied over
    For runQueryRow = 2 To totalSystems
        For assetSystemLineItemRow = 2 To assetSystemLineItemList
            If runQuerySheet.Cells(runQueryRow, 1).Value = assetSystemLineItems.Cells(assetSystemLineItemRow, 2) Then
                runQuerySheet.Cells(runQueryRow, 1).Interior.Color = 65535
            End If
        Next
    Next
    


End Sub

Function systemNameLink(row)
    
    systemName = Workbooks("RunQuery").Sheets("RunQuery").Cells(row, 2)
    findDash = Application.WorksheetFunction.find("-", systemName) - 2
    systemNameLink = Left(systemName, findDash)

End Function

Function numLineItem(row)
    
    ' returns number of line items
    numLineItem = Application.WorksheetFunction.CountA(Sheets("MASTER").Range("J" & row, "IM" & row)) / 7
        
End Function

Function masterLineItem(row, lineItem)

    masterLineItem = Sheets("MASTER").Range(Cells(row, (10 + (lineItem - 1) * 7)), Cells(row, (16 + (lineItem - 1) * 7))).Value

End Function

Function emptyCellRow(column)

    ' returns next empty row
    emptyCellRow = Sheets("Asset System Line Items").Cells(Rows.Count, column).End(xlUp).Offset(1, 0).row

End Function




