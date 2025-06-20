Sub GenerateMonthlySummaryOrdered()

    Dim wsSource As Worksheet, wsReport As Worksheet
    Dim lastRow As Long, reportRow As Long
    Dim dataDict As Object, mapDict As Object
    Dim excludeFundDict As Object, excludeAccountDict As Object
    Dim fundOrderDict As Object, mapAccountDict As Object
    Dim i As Long, j As Long
    Dim key As Variant
    Dim fund As String, mappedFund As String, desc As String
    Dim parent As String, adjParent As String, fiscalYear As String, account As String
    Dim periodTotal As Double
    Dim monthVal As Variant, monthNum As Integer, monthName As String
    Dim monthOrder As Variant, monthDict As Object
    Dim headers As Collection, allKeys As Collection

    ' Rename sheets
    On Error Resume Next
    ThisWorkbook.Sheets("Sheet1").Name = "Data"
    ThisWorkbook.Sheets("Sheet2").Name = "Revenue Report"
    ThisWorkbook.Sheets("Order").Name = "FundOrder"
    On Error GoTo 0

    ' Set sheets
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

    ' Load Fund Order
    Dim wsOrder As Worksheet: Set wsOrder = ThisWorkbook.Sheets("FundOrder")
    i = 1
    Do While wsOrder.Cells(i, 1).Value <> ""
        fundOrderDict(Trim(wsOrder.Cells(i, 1).Value)) = i
        i = i + 1
    Loop

    monthOrder = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' Process source data
    For i = 2 To lastRow
        fund = Trim(wsSource.Cells(i, "I").Text)
        parent = Trim(wsSource.Cells(i, "C").Text)
        account = Trim(wsSource.Cells(i, "E").Text)

        adjParent = IIf(Len(parent) > 1, Mid(parent, 2) & "00", "")
        If mapAccountDict.exists(fund & "|" & parent) Then
            adjParent = mapAccountDict(fund & "|" & parent)
        End If
        If Len(adjParent) > 6 Then adjParent = Left(adjParent, 6)

        If mapDict.exists(fund) Then
            mappedFund = mapDict(fund)
        Else
            mappedFund = fund
        End If

        If excludeFundDict.exists(mappedFund) Then GoTo SkipRow
        If excludeAccountDict.exists(account) Then GoTo SkipRow

        desc = Trim(wsSource.Cells(i, "D").Value)
        fiscalYear = Trim(wsSource.Cells(i, "A").Value)
        periodTotal = Val(wsSource.Cells(i, "G").Value)

        monthVal = wsSource.Cells(i, "B").Value
        If Not IsDate(monthVal) Then GoTo SkipRow
        monthNum = Month(monthVal)
        monthName = Format(DateSerial(1900, monthNum, 1), "mmm")
        monthDict(monthName) = True

        key = mappedFund & "|" & desc & "|" & parent & "|" & fiscalYear
        If Not dataDict.exists(key) Then
            Set dataDict(key) = CreateObject("Scripting.Dictionary")
            dataDict(key)("Fund") = mappedFund
            dataDict(key)("Description") = desc
            dataDict(key)("AdjustedParent") = adjParent
            dataDict(key)("FY") = fiscalYear
            dataDict(key)("Parent") = parent
            dataDict(key)("Total") = 0
            allKeys.Add key
        End If

        If dataDict(key).exists(monthName) Then
            dataDict(key)(monthName) = dataDict(key)(monthName) + periodTotal
        Else
            dataDict(key)(monthName) = periodTotal
        End If
        dataDict(key)("Total") = dataDict(key)("Total") + periodTotal

SkipRow:
    Next i

    ' Prepare headers
    Set headers = New Collection
    headers.Add "Fund"
    headers.Add "Description"
    headers.Add "SCO Revenue Code"
    For Each m In monthOrder
        If monthDict.exists(m) Then headers.Add m
    Next m
    headers.Add "Total"
    headers.Add "FY"

    ' Write headers
    For i = 1 To headers.Count
        wsReport.Cells(1, i).Value = headers(i)
        wsReport.Cells(1, i).Font.Bold = True
    Next i

    ' Sort keys by Fund Order
    Dim sortedKeys() As String
    ReDim sortedKeys(1 To allKeys.Count)
    For i = 1 To allKeys.Count
        sortedKeys(i) = allKeys(i)
    Next i

    For i = 1 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            Dim keyA As Variant: keyA = Split(sortedKeys(i), "|")
            Dim keyB As Variant: keyB = Split(sortedKeys(j), "|")
            Dim fundA As String: fundA = keyA(0)
            Dim fundB As String: fundB = keyB(0)
            Dim orderA As Long: orderA = IIf(fundOrderDict.exists(fundA), fundOrderDict(fundA), 999999)
            Dim orderB As Long: orderB = IIf(fundOrderDict.exists(fundB), fundOrderDict(fundB), 999999)

            If orderA > orderB Or (orderA = orderB And fundA > fundB) Then
                Dim temp As String
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Write data rows
    reportRow = 2
    Dim monthColStart As Integer: monthColStart = 4
    Dim totalColIndex As Integer

    For i = 1 To UBound(sortedKeys)
        key = sortedKeys(i)
        Dim dict As Object: Set dict = dataDict(key)
        Dim col As Integer: col = 1

        wsReport.Cells(reportRow, col).Value = "'" & dict("Fund"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("Description"): col = col + 1
        wsReport.Cells(reportRow, col).Value = "'" & dict("AdjustedParent"): col = col + 1

        For Each m In monthOrder
            If monthDict.exists(m) Then
                wsReport.Cells(reportRow, col).Value = IIf(dict.exists(m), dict(m), 0)
                col = col + 1
            End If
        Next m

        totalColIndex = col
        wsReport.Cells(reportRow, totalColIndex).Value = dict("Total")
        wsReport.Cells(reportRow, totalColIndex).Font.Bold = True
        col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("FY")

        ' Highlight yellow if Fund = 094001
        If dict("Fund") = "094001" Then
            wsReport.Range(wsReport.Cells(reportRow, 1), wsReport.Cells(reportRow, headers.Count)).Interior.Color = RGB(255, 255, 153)
        End If

        ' Highlight red if negative total
        If dict("Total") < 0 Then
            wsReport.Range(wsReport.Cells(reportRow, 1), wsReport.Cells(reportRow, headers.Count)).Interior.Color = RGB(255, 199, 206)
        End If

        reportRow = reportRow + 1
    Next i

    ' Format numbers
    Dim colIndex As Integer
    For colIndex = monthColStart To monthColStart + monthDict.Count
        wsReport.Columns(colIndex).NumberFormat = "#,##0.00"
    Next colIndex

    wsReport.Columns.AutoFit
    MsgBox "Revenue Report generated successfully.", vbInformation

End Sub
