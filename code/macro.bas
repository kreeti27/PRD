Sub GenerateMonthlySummaryOrdered()
    Dim wsSource As Worksheet, wsReport As Worksheet
    Dim lastRow As Long, reportRow As Long
    Dim dataDict As Object, mapDict As Object
    Dim excludeFundDict As Object, excludeAccountDict As Object
    Dim fundOrderDict As Object, mapAccountDict As Object
    Dim i As Long, key As Variant
    Dim fund As String, mappedFund As String, desc As String, parent As String, adjParent As String
    Dim fiscalYear As String, account As String
    Dim periodTotal As Double
    Dim monthVal As Variant, monthNum As Integer, monthName As String
    Dim monthOrder As Variant, monthDict As Object
    Dim headers As Collection, allKeys As Collection

    ' Rename sheets if necessary
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    If Not wsSource Is Nothing Then wsSource.Name = "Data"
    Set wsReport = ThisWorkbook.Sheets("Sheet2")
    If Not wsReport Is Nothing Then wsReport.Name = "Revenue Report"
    On Error GoTo 0

    Set wsSource = ThisWorkbook.Sheets("Data")
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Revenue Report")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=wsSource)
        wsReport.Name = "Revenue Report"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    Set dataDict = CreateObject("Scripting.Dictionary")
    Set monthDict = CreateObject("Scripting.Dictionary")
    Set mapDict = CreateObject("Scripting.Dictionary")
    Set excludeFundDict = CreateObject("Scripting.Dictionary")
    Set excludeAccountDict = CreateObject("Scripting.Dictionary")
    Set fundOrderDict = CreateObject("Scripting.Dictionary")
    Set mapAccountDict = CreateObject("Scripting.Dictionary")
    Set allKeys = New Collection

    ' Load MappingAccount
    Dim wsMapAcc As Worksheet: Set wsMapAcc = ThisWorkbook.Sheets("MappingAccount")
    i = 2
    Do While wsMapAcc.Cells(i, 1).Value <> ""
        mapAccountDict(Trim(wsMapAcc.Cells(i, 1).Value) & "|" & Trim(wsMapAcc.Cells(i, 2).Value)) = Trim(wsMapAcc.Cells(i, 3).Value)
        i = i + 1
    Loop

    ' Load MappingFund
    Dim wsMap As Worksheet: Set wsMap = ThisWorkbook.Sheets("MappingFund")
    i = 1
    Do While wsMap.Cells(i, 1).Value <> ""
        mapDict(Trim(wsMap.Cells(i, 1).Value)) = Trim(wsMap.Cells(i, 2).Value)
        i = i + 1
    Loop

    ' Load ExcludeFund
    Dim wsExclude As Worksheet: Set wsExclude = ThisWorkbook.Sheets("ExcludeFund")
    i = 1
    Do While wsExclude.Cells(i, 1).Value <> ""
        excludeFundDict(Trim(wsExclude.Cells(i, 1).Value)) = True
        i = i + 1
    Loop

    ' Load ExcludeAccounts
    Dim wsExcludeAcc As Worksheet: Set wsExcludeAcc = ThisWorkbook.Sheets("ExcludeAccounts")
    i = 1
    Do While wsExcludeAcc.Cells(i, 1).Value <> ""
        excludeAccountDict(Trim(wsExcludeAcc.Cells(i, 1).Value)) = True
        i = i + 1
    Loop

    ' Load FundOrder
    Dim wsOrder As Worksheet: Set wsOrder = ThisWorkbook.Sheets("FundOrder")
    i = 1
    Do While wsOrder.Cells(i, 1).Value <> ""
        fundOrderDict(Trim(wsOrder.Cells(i, 1).Value)) = i
        i = i + 1
    Loop

    ' Define month order
    monthOrder = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' Process rows
    For i = 2 To lastRow
        fund = Trim(wsSource.Cells(i, "I").Text)
        parent = Trim(wsSource.Cells(i, "C").Text)
        account = Trim(wsSource.Cells(i, "E").Text)

        If mapDict.exists(fund) Then
            mappedFund = mapDict(fund)
        Else
            mappedFund = fund
        End If

        If excludeFundDict.exists(mappedFund) Then GoTo SkipRow
        If excludeAccountDict.exists(account) Then GoTo SkipRow

        ' Default Adjusted Parent
        adjParent = ""
        If Len(parent) > 1 Then adjParent = Mid(parent, 2) & "00"
        If mapAccountDict.exists(fund & "|" & parent) Then
            adjParent = mapAccountDict(fund & "|" & parent)
        End If

        ' Accept leading zeros and ensure 6 digits max from right
        If Len(adjParent) > 6 Then
            adjParent = Right(adjParent, 6)
        ElseIf Len(adjParent) < 6 Then
            adjParent = Right("000000" & adjParent, 6)
        End If

        desc = Trim(wsSource.Cells(i, "D").Value)
        fiscalYear = Trim(wsSource.Cells(i, "A").Value)
        periodTotal = Val(wsSource.Cells(i, "G").Value)

        monthVal = wsSource.Cells(i, "B").Value
        If Not IsDate(monthVal) Then GoTo SkipRow
        monthNum = Month(monthVal)
        monthName = Format(DateSerial(1900, monthNum, 1), "mmm")
        monthDict(monthName) = True

        key = fiscalYear & "|" & mappedFund & "|" & desc & "|" & adjParent
        If Not dataDict.exists(key) Then
            Set dataDict(key) = CreateObject("Scripting.Dictionary")
            dataDict(key)("FY") = fiscalYear
            dataDict(key)("Fund") = mappedFund
            dataDict(key)("Description") = desc
            dataDict(key)("SCOCode") = adjParent
            Set dataDict(key)("Months") = CreateObject("Scripting.Dictionary")
            allKeys.Add key
        End If

        If dataDict(key)("Months").exists(monthName) Then
            dataDict(key)("Months")(monthName) = dataDict(key)("Months")(monthName) + periodTotal
        Else
            dataDict(key)("Months")(monthName) = periodTotal
        End If

SkipRow:
    Next i

    ' Build headers
    Set headers = New Collection
    headers.Add "Fund"
    headers.Add "Description"
    headers.Add "SCO Revenue Code"
    For Each m In monthOrder
        If monthDict.exists(m) Then headers.Add m
    Next m
    headers.Add "FY"

    ' Write headers
    For i = 1 To headers.Count
        wsReport.Cells(1, i).Value = headers(i)
        wsReport.Cells(1, i).Font.Bold = True
    Next i

    ' Sort keys by FY, fundOrder, SCO Revenue Code
    Dim sortedKeys() As String
    ReDim sortedKeys(1 To allKeys.Count)
    For i = 1 To allKeys.Count
        sortedKeys(i) = allKeys(i)
    Next i

    Dim j As Long
    For i = 1 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            Dim a(): a = Split(sortedKeys(i), "|")
            Dim b(): b = Split(sortedKeys(j), "|")

            Dim fyA As Long: fyA = Val(a(0))
            Dim fyB As Long: fyB = Val(b(0))

            Dim fundA As String: fundA = a(1)
            Dim fundB As String: fundB = b(1)

            Dim scoA As Long: scoA = Val(a(3))
            Dim scoB As Long: scoB = Val(b(3))

            Dim orderA As Long: orderA = IIf(fundOrderDict.exists(fundA), fundOrderDict(fundA), 999999)
            Dim orderB As Long: orderB = IIf(fundOrderDict.exists(fundB), fundOrderDict(fundB), 999999)

            If fyA > fyB _
                Or (fyA = fyB And orderA > orderB) _
                Or (fyA = fyB And orderA = orderB And scoA > scoB) Then
                Dim temp As String
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Write data
    reportRow = 2
    For i = 1 To UBound(sortedKeys)
        key = sortedKeys(i)
        Dim d As Object: Set d = dataDict(key)
        Dim col As Integer: col = 1

        wsReport.Cells(reportRow, col).Value = "'" & d("Fund"): col = col + 1
        wsReport.Cells(reportRow, col).Value = d("Description"): col = col + 1
        wsReport.Cells(reportRow, col).Value = d("SCOCode"): col = col + 1

        For Each m In monthOrder
            If monthDict.exists(m) Then
                If d("Months").exists(m) Then
                    wsReport.Cells(reportRow, col).Value = d("Months")(m)
                Else
                    wsReport.Cells(reportRow, col).Value = 0
                End If
                col = col + 1
            End If
        Next m

        wsReport.Cells(reportRow, col).Value = d("FY")

        ' Highlight specific fund
        If d("Fund") = "0044094" Then
            wsReport.Range(wsReport.Cells(reportRow, 1), wsReport.Cells(reportRow, headers.Count)).Interior.Color = RGB(255, 255, 153)
        End If

        reportRow = reportRow + 1
    Next i

    ' Format month columns
    Dim colIndex As Integer, startCol As Integer
    startCol = 4
    For colIndex = startCol To startCol + monthDict.Count - 1
        wsReport.Columns(colIndex).NumberFormat = "#,##0.00"
    Next colIndex

    wsReport.Columns.AutoFit
    MsgBox "Report generated successfully in 'Revenue Report'.", vbInformation
End Sub
