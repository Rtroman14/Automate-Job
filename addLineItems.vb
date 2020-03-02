Sub addLineItems()
'

    ' MASTER TEMPLATE variables
    Dim masterTemplate As Workbook
    Dim MASTER As Worksheet
    Set masterTemplate = Workbooks("MASTER TEMPLATE.xlsm")
    Set MASTER = masterTemplate.Worksheets("MASTER")

    ' ReportDownload variables
    Dim reportDownload As Workbook
    Set reportDownload = Workbooks("ReportDownload")

    ' variables
    Dim ws As Worksheet
    Dim totalSystems As Integer
    Dim row As Integer
    Dim masterSystemRow As Integer
    Dim lineItemHeading As Range
    Dim lineItemRow As Integer
    Dim totalLineItems As Integer
    Dim lineItem As Integer
    Dim lineItemCol As Integer
    Dim lineItemCell As Integer
    
    

    ' total systems in MASTER
    totalSystems = Application.CountA(MASTER.Range("A4:A1000"))

    ' loop through each worksheet in Workbook "ReportDownload"
    For Each ws In reportDownload.Worksheets
    
        ' locate system name and save as variable
        newSystem = ws.Range("D10").Value
        findDash = Application.WorksheetFunction.find("-", newSystem) - 2
        newSystemName = Left(newSystem, findDash)
        
        Debug.Print newSystemName


        ' loop through MASTER TEMPLATE to see if system exists
        For row = 1 To totalSystems

            If MASTER.Cells(row + 3, 1).Value = newSystemName Then
            
            ' return row number of located system
                masterSystemRow = row + 3

            ' search column "E" in "ReportDownload" for text "Quantity"
                Set lineItemHeading = ws.Range("E:E").find("Quantity")

                ' lineItemRow = (row number + 1)
                lineItemRow = lineItemHeading.row + 1

                ' count number of line items
                totalLineItems = Application.CountA(ws.Range("F:F")) - 1

                ' increment column variable
                lineItemCol = 10
                
                'loop through lineItemCount
                For lineItem = 1 To totalLineItems

                    ' check if column ("F" + lineItemRow) > 1
                    If ws.Cells(lineItemRow - 1 + lineItem, 5) > 1 Then

                        ' loop through each cell in line item row and transfer to MASTER
                        For lineItemCell = 1 To 8

                            If Not lineItemCell = 6 Then

                                MASTER.Cells(masterSystemRow, lineItemCol) = ws.Cells(lineItemRow - 1 + lineItem, lineItemCell)

                                lineItemCol = lineItemCol + 1

                            End If

                        Next
                    
                    End If

                Next

            End If
                        
        Next

    Next ws
    

End Sub

