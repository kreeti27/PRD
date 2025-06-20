Sub GenerateMonthlySummaryOrdered()

    ' Declare worksheet variables
    Dim wsSource As Worksheet, wsReport As Worksheet
    Dim wsMapAcc As Worksheet, wsMap As Worksheet
    Dim wsExclude As Worksheet, wsExcludeAcc As Worksheet, wsOrder As Worksheet

    ' Declare collections and dictionaries
    Dim dataDict As Object, mapDict As Object
    Dim excludeFundDict As Object, excludeAccountDict As Object
    Dim fundOrderDict As Object, mapAccountDict As Object
    Dim monthDict As Object
    Dim headers As Collection, allKeys As Collection

    ' Declare general variables
    Dim lastRow As Long, reportRow As Long, i As Long
    Dim fund As String, mappedFund As String, desc As String
    Dim parent As String, adjParent As String, fiscalYear As String, account As String
    Dim periodTotal As Double, monthVal As Variant, monthNum As Integer, monthName As String
    Dim monthOrder As Variant, key As Variant
    Dim m As Variant, k As Variant

    ' Set worksheet references
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Sheet2")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=wsSource)
        wsReport.Name = "Sheet2"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0

    ' Initialize dictionaries and collections
    Set dataDict = CreateObject("Scripting.Dictionary")
    Set monthDict = CreateObject("Scripting.Dictionary")
    Set mapDict = CreateObject("Scripting.Dictionary")
    Set excludeFundDict = CreateObject("Scripting.Dictionary")
    Set excludeAccountDict = CreateObject("Scripting.Dictionary")
    Set fundOrderDict = CreateObject("Scripting.Dictionary")
    Set mapAccountDict = CreateObject("Scripting.Dictionary")
    Set allKeys = New Collection

    ' Load mapping sheets
    Set wsMapAcc = ThisWorkbook.Sheets("MappingAccount")
    i = 2
    Do While wsMapAcc.Cells(i, 1).Value <> ""
        mapAccountDict(Trim(wsMapAcc.Cells(i, 1).Value) & "|" & Trim(wsMapAcc.Cells(i, 2).Value)) = Trim(wsMapAcc.Cells(i, 3).Value)
        i = i + 1
    Loop

    Set wsMap = ThisWorkbook.Sheets("MappingFund")
    i = 1
    Do While wsMap.Cells(i, 1).Value <> ""
        mapDict(Trim(wsMap.Cells(i, 1).Value)) = Trim(wsMap.Cells(i, 2).Value)
        i = i + 1
    Loop

    Set wsExclude = ThisWorkbook.Sheets("ExcludeFund")
    i = 1
    Do While wsExclude.Cells(i, 1).Value <> ""
        excludeFundDict(Trim(wsExclude.Cells(i, 1).Value)) = True
        i = i + 1
    Loop

    Set wsExcludeAcc = ThisWorkbook.Sheets("ExcludeAccounts")
    i = 1
    Do While wsExcludeAcc.Cells(i, 1).Value <> ""
        excludeAccountDict(Trim(wsExcludeAcc.Cells(i, 1).Value)) = True
        i = i + 1
    Loop

    Set wsOrder = ThisWorkbook.Sheets("Order")
    i = 1
    Do While wsOrder.Cells(i, 1).Value <> ""
        fundOrderDict(Trim(wsOrder.Cells(i, 1).Value)) = i
        i = i + 1
    Loop

    ' Month order used for output
    monthOrder = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' Read and process source data
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        fund = Trim(wsSource.Cells(i, "I").Text)
        parent = Trim(wsSource.Cells(i, "C").Text)
        account = Trim(wsSource.Cells(i, "E").Text)

        ' Set default adjusted parent
        adjParent = IIf(Len(parent) > 1, Mid(parent, 2) & "00", "")
        If mapAccountDict.exists(fund & "|" & parent) Then
            adjParent = mapAccountDict(fund & "|" & parent)
        End If

        ' Apply fund mapping
        If mapDict.exists(fund) Then
            mappedFund = mapDict(fund)
        Else
            mappedFund = fund
        End If

        ' Skip if fund or account is excluded
        If excludeFundDict.exists(mappedFund) Or excludeAccountDict.exists(account) Then GoTo SkipRow

        ' Collect remaining fields
        desc = Trim(wsSource.Cells(i, "D").Value)
        fiscalYear = Trim(wsSource.Cells(i, "A").Value)
        periodTotal = Val(wsSource.Cells(i, "G").Value)
        monthVal = wsSource.Cells(i, "B").Value
        If Not IsDate(monthVal) Then GoTo SkipRow

        monthNum = Month(monthVal)
        monthName = Format(DateSerial(1900, monthNum, 1), "mmm")
        monthDict(monthName) = True

        ' Construct aggregation key
        key = mappedFund & "|" & desc & "|" & parent & "|" & fiscalYear

        ' Create entry if it does not exist
        If Not dataDict.exists(key) Then
            Set dataDict(key) = CreateObject("Scripting.Dictionary")
            dataDict(key)("Fund") = mappedFund
            dataDict(key)("Description") = desc
            dataDict(key)("Parent") = parent
            dataDict(key)("AdjustedParent") = adjParent
            dataDict(key)("FY") = fiscalYear
            dataDict(key)("Total") = 0
            allKeys.Add key
        End If

        ' Aggregate by month and total
        If dataDict(key).exists(monthName) Then
            dataDict(key)(monthName) = dataDict(key)(monthName) + periodTotal
        Else
            dataDict(key)(monthName) = periodTotal
        End If
        dataDict(key)("Total") = dataDict(key)("Total") + periodTotal

SkipRow:
    Next i

    ' Prepare header list
    Set headers = New Collection
    headers.Add "Fund"
    headers.Add "Description"
    headers.Add "Parent Code"
    headers.Add "Adjusted Parent"
    For Each m In monthOrder
        If monthDict.exists(m) Then headers.Add m
    Next m
    headers.Add "Total"
    headers.Add "FY"

    ' Write headers to report
    For i = 1 To headers.Count
        wsReport.Cells(1, i).Value = headers(i)
        wsReport.Cells(1, i).Font.Bold = True
    Next i

    ' Sort keys based on custom fund order and parent code
    Dim sortedKeys() As String
    ReDim sortedKeys(1 To allKeys.Count)
    For i = 1 To allKeys.Count
        sortedKeys(i) = allKeys(i)
    Next i

    Dim j As Long
    For i = 1 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            Dim keyA As Variant: keyA = Split(sortedKeys(i), "|")
            Dim keyB As Variant: keyB = Split(sortedKeys(j), "|")
            Dim fundA As String: fundA = keyA(0)
            Dim fundB As String: fundB = keyB(0)
            Dim orderA As Long: orderA = IIf(fundOrderDict.exists(fundA), fundOrderDict(fundA), 999999)
            Dim orderB As Long: orderB = IIf(fundOrderDict.exists(fundB), fundOrderDict(fundB), 999999)
            Dim parentA As Long, parentB As Long
            parentA = IIf(IsNumeric(keyA(2)), CLng(keyA(2)), 9999999)
            parentB = IIf(IsNumeric(keyB(2)), CLng(keyB(2)), 9999999)

            If orderA > orderB Or _
               (orderA = orderB And fundA > fundB) Or _
               (orderA = orderB And fundA = fundB And parentA > parentB) Then
                Dim temp As String
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Write sorted data to report
    reportRow = 2
    Dim monthColStart As Integer: monthColStart = 5
    Dim totalColIndex As Integer, col As Integer
    Dim dict As Object

    For i = 1 To UBound(sortedKeys)
        key = sortedKeys(i)
        Set dict = dataDict(key)
        col = 1

        wsReport.Cells(reportRow, col).Value = "'" & dict("Fund"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("Description"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("Parent"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("AdjustedParent"): col = col + 1

        For Each m In monthOrder
            If monthDict.exists(m) Then
                wsReport.Cells(reportRow, col).Value = IIf(dict.exists(m), dict(m), 0)
                col = col + 1
            End If
        Next m

        totalColIndex = col
        wsReport.Cells(reportRow, totalColIndex).Value = dict("Total"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("FY")

        If dict("Total") < 0 Then
            wsReport.Range(wsReport.Cells(reportRow, 1), wsReport.Cells(reportRow, headers.Count)).Interior.Color = RGB(255, 199, 206)
        End If

        reportRow = reportRow + 1
    Next i

    ' Apply number formatting to month and total columns
    Dim colIndex As Integer
    For colIndex = monthColStart To monthColStart + monthDict.Count
        wsReport.Columns(colIndex).NumberFormat = "#,##0.00"
    Next colIndex

    wsReport.Columns.AutoFit
    MsgBox "Ordered report generated in Sheet2 with MappingAccount and MappingFund applied.", vbInformation

End Sub
