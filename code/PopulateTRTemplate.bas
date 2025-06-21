Sub PopulateTRTemplate()

    Dim wsSource As Worksheet, wsDest As Worksheet, wsMap As Worksheet
    Dim lastRow As Long, destRow As Long
    Dim fundMap As Object
    Dim i As Long
    Dim transferReq As String, fiscalYear As String
    Dim fund As String, scoCode As String, totalVal As Double, fy As String
    Dim agency As String, account As String, colA As String
    Dim amountSum As Double

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

            ' Preserve leading zeros by forcing text with apostrophe
            fund = "'" & fund
            scoCode = "'" & scoCode
            If agency <> "" Then agency = "'" & agency

            With wsDest
                .Cells(destRow, 1).Value = fund                 ' Fund
                .Cells(destRow, 2).Value = agency               ' Agency
                .Cells(destRow, 3).Value = fiscalYear           ' Fiscal Year
                .Cells(destRow, 4).Value = ""                   ' Ref Item
                .Cells(destRow, 5).Value = ""                   ' Fed Cat
                .Cells(destRow, 6).Value = ""                   ' P/N
                .Cells(destRow, 7).Value = ""                   ' C
                .Cells(destRow, 8).Value = ""                   ' Cat
                .Cells(destRow, 9).Value = ""                   ' Pgm
                .Cells(destRow, 10).Value = ""                  ' Ele
                .Cells(destRow, 11).Value = ""                  ' Comp
                .Cells(destRow, 12).Value = ""                  ' Task
                .Cells(destRow, 13).Value = account             ' Account
                .Cells(destRow, 14).Value = scoCode             ' Rev/Obj
                .Cells(destRow, 15).Value = "C"                 ' D/C
                .Cells(destRow, 16).Value = colA                ' A
                .Cells(destRow, 17).Value = ""                  ' Source Fund
                .Cells(destRow, 18).Value = totalVal            ' Amount
                .Cells(destRow, 18).NumberFormat = "#,##0.00"
                .Cells(destRow, 19).Value = "TRF REQ " & transferReq  ' Description
                .Cells(destRow, 20).Value = ""                  ' DNKP
                .Cells(destRow, 21).Value = ""                  ' Prgm Desc
            End With

            amountSum = amountSum + totalVal
            destRow = destRow + 1
        End If
    Next i

    ' Set R2 and S2
    wsDest.Range("R2").Value = amountSum
    wsDest.Range("R2").NumberFormat = "#,##0.00"
    wsDest.Range("S2").Value = "TRF REQ " & transferReq

    MsgBox "TR Template has been populated successfully.", vbInformation

End Sub
