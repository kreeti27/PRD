Sub PopulateTRTemplate()

    Dim wsSource As Worksheet, wsDest As Worksheet, wsMap As Worksheet
    Dim lastRow As Long, destRow As Long
    Dim fundMap As Object
    Dim i As Long
    Dim transferReq As String, fiscalYear As String
    Dim fund As String, scoCode As String, totalVal As Double, fy As String
    Dim agency As String, account As String, colA As String
    Dim amountSum As Double
    Dim savePath As String, saveFileName As String

    ' Prompt for inputs
    transferReq = InputBox("Enter Transfer Request Number")
    If transferReq = "" Then Exit Sub

    fiscalYear = InputBox("Enter Fiscal Year")
    If fiscalYear = "" Then Exit Sub

    ' Set worksheets
    Set wsSource = Sheets("Revenue Report")
    Set wsDest = Sheets("TR Template")
    Set wsMap = Sheets("AgencyMapping")

    ' Validate if Revenue Report has data
    If Application.WorksheetFunction.CountA(wsSource.Cells) = 0 Then
        MsgBox "Revenue Report is empty.", vbExclamation
        Exit Sub
    End If

    ' Load mapping dictionary
    Set fundMap = CreateObject("Scripting.Dictionary")
    Dim mapLastRow As Long
    mapLastRow = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row
    For i = 2 To mapLastRow
        fundMap(wsMap.Cells(i, 1).Text) = wsMap.Cells(i, 2).Text
    Next i

    ' Clear previous data from A3 down and R2, S2
    wsDest.Range("A3:U" & wsDest.Rows.Count).ClearContents
    wsDest.Range("R2").ClearContents
    wsDest.Range("S2").ClearContents

    ' Find last row in Revenue Report
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    destRow = 3
    amountSum = 0

    ' Find column indexes
    Dim colFund As Integer, colCode As Integer, colTotal As Integer, colFY As Integer
    colFund = Application.Match("Fund", wsSource.Rows(1), 0)
    colCode = Application.Match("SCO Revenue Code", wsSource.Rows(1), 0)
    colTotal = Application.Match("Total", wsSource.Rows(1), 0)
    colFY = Application.Match("FY", wsSource.Rows(1), 0)

    For i = 2 To lastRow
        If Trim(wsSource.Cells(i, colFY).Text) = fiscalYear Then
            fund = Trim(wsSource.Cells(i, colFund).Text)
            scoCode = Trim(wsSource.Cells(i, colCode).Text)
            totalVal = wsSource.Cells(i, colTotal).Value

            ' Default values
            account = ""
            colA = "R"

            ' If code is 084000, override fund, account, A, and agency
            If scoCode = "084000" Then
                fund = "0044"
                account = "3730"
                colA = "G"
                agency = ""  ' override with blank
            Else
                If fundMap.exists(fund) Then
                    agency = fundMap(fund)
                Else
                    agency = ""
                End If
            End If

            ' Preserve leading zeros by forcing text
            fund = "'" & fund
            scoCode = "'" & scoCode
            If agency <> "" Then agency = "'" & agency

            With wsDest
                .Cells(destRow, 1).Value = fund
                .Cells(destRow, 2).Value = agency
                .Cells(destRow, 3).Value = fiscalYear
                .Cells(destRow, 4).Value = ""
                .Cells(destRow, 5).Value = ""
                .Cells(destRow, 6).Value = ""
                .Cells(destRow, 7).Value = ""
                .Cells(destRow, 8).Value = ""
                .Cells(destRow, 9).Value = ""
                .Cells(destRow, 10).Value = ""
                .Cells(destRow, 11).Value = ""
                .Cells(destRow, 12).Value = ""
                .Cells(destRow, 13).Value = account
                .Cells(destRow, 14).Value = scoCode
                .Cells(destRow, 15).Value = "C"
                .Cells(destRow, 16).Value = colA
                .Cells(destRow, 17).Value = ""
                .Cells(destRow, 18).Value = totalVal
                .Cells(destRow, 18).NumberFormat = "#,##0.00"
                .Cells(destRow, 19).Value = "TRF REQ " & transferReq
                .Cells(destRow, 20).Value = ""
                .Cells(destRow, 21).Value = ""
            End With

            amountSum = amountSum + totalVal
            destRow = destRow + 1
        End If
    Next i

    ' Set R2 and S2
    wsDest.Range("R2").Value = amountSum
    wsDest.Range("R2").NumberFormat = "#,##0.00"
    wsDest.Range("S2").Value = "TRF REQ " & transferReq

    ' Ask user for location to save
    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

    If fDialog.Show <> -1 Then
        MsgBox "Save cancelled. The workbook was not saved.", vbExclamation
        Exit Sub
    End If

    savePath = fDialog.SelectedItems(1)
    saveFileName = "TR" & transferReq & "_" & Format(Date, "MM-dd-yy") & ".xlsx"

    ' Save copy of current workbook
    ThisWorkbook.SaveCopyAs savePath & "\" & saveFileName

    MsgBox "TR Template has been populated and saved successfully as '" & saveFileName & "'.", vbInformation

End Sub
