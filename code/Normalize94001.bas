Sub Normalize0094001Rows()

    Dim wsRev As Worksheet, wsCode As Worksheet

    Dim lastRow As Long, matchRows As Collection

    Dim i As Long, scoCol As Long

    Dim newVal As Variant

   

    Set wsRev = ThisWorkbook.Sheets("Revenue Report")

    Set wsCode = ThisWorkbook.Sheets("0094001 Revenue Code")

   

    Set matchRows = New Collection

   

    ' Find matching rows with 0094001 in column A

    lastRow = wsRev.Cells(wsRev.Rows.Count, "A").End(xlUp).Row

    For i = 1 To lastRow

        If wsRev.Cells(i, 1).Text = "0094001" Then

            matchRows.Add i

        End If

    Next i

   

    ' Get the column number of "SCO Revenue code"

    On Error Resume Next

    scoCol = Application.WorksheetFunction.Match("SCO Revenue code", wsRev.Rows(1), 0)

    On Error GoTo 0

   

    If scoCol = 0 Then

        MsgBox "Column 'SCO Revenue code' not found!", vbExclamation

        Exit Sub

    End If

   

    ' Skip the whole normalize routine if there are no records for 0094001

    If matchRows.Count > 0 Then

        ' If fewer than 5, duplicate the last one

        If matchRows.Count < 5 And matchRows.Count > 0 Then

            Dim lastMatchRow As Long

            lastMatchRow = matchRows(matchRows.Count)

            For i = matchRows.Count + 1 To 5

                wsRev.Rows(lastMatchRow + 1).Insert Shift:=xlDown

                wsRev.Rows(lastMatchRow).Copy wsRev.Rows(lastMatchRow + 1)

                matchRows.Add lastMatchRow + 1

                lastMatchRow = lastMatchRow + 1

            Next i

        ElseIf matchRows.Count > 5 Then

            ' Delete extra rows from the end

            For i = matchRows.Count To 6 Step -1

                wsRev.Rows(matchRows(i)).Delete

            Next i

            ' Rebuild matchRows collection

            Set matchRows = New Collection

            lastRow = wsRev.Cells(wsRev.Rows.Count, "A").End(xlUp).Row

            For i = 1 To lastRow

                If wsRev.Cells(i, 1).Text = "0094001" Then

                    matchRows.Add i

                End If

            Next i

        End If

       

        ' Override SCO Revenue code with values from 0094001 Revenue Code sheet

        fyCol = Application.WorksheetFunction.Match("FY", wsRev.Rows(1), 0)

        For i = 1 To 5

            newVal = wsCode.Cells(i, 1).Value

            wsRev.Cells(matchRows(i), scoCol).Value = "'" & newVal

            ' loop through all the columns after scoCol to fyCol-1 and make them blank

            For j = scoCol + 1 To fyCol - 1

                wsRev.Cells(matchRows(i), j).Value = ""

            Next j

        Next i

    End If

End Sub