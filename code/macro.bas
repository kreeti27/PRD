Sub GenerateMonthlySummaryOrdered()

    Dim wsSource As Worksheet, wsReport As Worksheet
    Dim lastRow As Long, reportRow As Long
    Dim dataDict As Object, mapDict As Object
    Dim excludeFundDict As Object, excludeAccountDict As Object
    Dim fundOrderDict As Object
    Dim mapAccountDict As Object
    Dim i As Long
    Dim key As Variant
    Dim fund As String, mappedFund As String, desc As String, parent As String, adjParent As String
    Dim fiscalYear As String, account As String
    Dim periodTotal As Double
    Dim monthVal As Variant, monthNum As Integer, monthName As String
    Dim monthOrder As Variant, monthDict As Object
    Dim headers As Collection
    Dim allKeys As Collection

    ' Set sheets
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

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    Set dataDict = CreateObject("Scripting.Dictionary")
    Set monthDict = CreateObject("Scripting.Dictionary")
    Set mapDict = CreateObject("Scripting.Dictionary")
    Set excludeFundDict = CreateObject("Scripting.Dictionary")
    Set excludeAccountDict = CreateObject("Scripting.Dictionary")
    Set fundOrderDict = CreateObject("Scripting.Dictionary")
    Set mapAccountDict = CreateObject("Scripting.Dictionary")
    Set allKeys = New Collection

    ' Load MappingAccount (with header)
    Dim wsMapAcc As Worksheet
    Set wsMapAcc = ThisWorkbook.Sheets("MappingAccount")
    i = 2
    Do While wsMapAcc.Cells(i, 1).Value <> ""
        ' Key = Fund|Parent, value = SCO Account
        mapAccountDict(Trim(wsMapAcc.Cells(i, 1).Value) & "|" & Trim(wsMapAcc.Cells(i, 2).Value)) = Trim(wsMapAcc.Cells(i, 3).Value)
        i = i + 1
    Loop

    ' Load MappingFund (no header)
    Dim wsMap As Worksheet: Set wsMap = ThisWorkbook.Sheets("MappingFund")
    i = 1
    Do While wsMap.Cells(i, 1).Value <> ""
        mapDict(Trim(wsMap.Cells(i, 1).Value)) = Trim(wsMap.Cells(i, 2).Value)
        i = i + 1
    Loop

    ' Load ExcludeFund (no header)
    Dim wsExclude As Worksheet: Set wsExclude = ThisWorkbook.Sheets("ExcludeFund")
    i = 1
    Do While wsExclude.Cells(i, 1).Value <> ""
        excludeFundDict(Trim(wsExclude.Cells(i, 1).Value)) = True
        i = i + 1
    Loop

    ' Load ExcludeAccounts (no header)
    Dim wsExcludeAcc As Worksheet: Set wsExcludeAcc = ThisWorkbook.Sheets("ExcludeAccounts")
    i = 1
    Do While wsExcludeAcc.Cells(i, 1).Value <> ""
        excludeAccountDict(Trim(wsExcludeAcc.Cells(i, 1).Value)) = True
        i = i + 1
    Loop

    ' Load Order sheet (no header)
    Dim wsOrder As Worksheet: Set wsOrder = ThisWorkbook.Sheets("Order")
    i = 1
    Do While wsOrder.Cells(i, 1).Value <> ""
        fundOrderDict(Trim(wsOrder.Cells(i, 1).Value)) = i
        i = i + 1
    Loop

    ' Month order
    monthOrder = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' Process source data
    For i = 2 To lastRow
        fund = Trim(wsSource.Cells(i, "I").Text)
        parent = Trim(wsSource.Cells(i, "C").Text)
        account = Trim(wsSource.Cells(i, "E").Text)

        ' Default Adjusted Parent as before
        adjParent = IIf(Len(parent) > 1, Mid(parent, 2) & "00", "")

        ' Override Adjusted Parent if MappingAccount has entry Fund|Parent
        If mapAccountDict.exists(fund & "|" & parent) Then
            adjParent = mapAccountDict(fund & "|" & parent)
        End If

        ' Apply MappingFund for Fund (only changes fund)
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
            dataDict(key)("Parent") = parent
            dataDict(key)("AdjustedParent") = adjParent
            dataDict(key)("FY") = fiscalYear
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
    headers.Add "Parent Code"
    headers.Add "Adjusted Parent"
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

    ' Sort keys
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
            If IsNumeric(keyA(2)) Then parentA = CLng(keyA(2)) Else parentA = 9999999
            If IsNumeric(keyB(2)) Then parentB = CLng(keyB(2)) Else parentB = 9999999

            If orderA > orderB _
                Or (orderA = orderB And fundA > fundB) _
                Or (orderA = orderB And fundA = fundB And parentA > parentB) Then
                Dim temp As String
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i

    ' Write data
    reportRow = 2
    Dim monthColStart As Integer: monthColStart = 5 ' After Fund, Desc, Parent, Adjusted Parent
    Dim totalColIndex As Integer

    For i = 1 To UBound(sortedKeys)
        key = sortedKeys(i)
        Dim dict As Object: Set dict = dataDict(key)
        Dim col As Integer: col = 1

        wsReport.Cells(reportRow, col).Value = "'" & dict("Fund"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("Description"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("Parent"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("AdjustedParent"): col = col + 1

        For Each m In monthOrder
            If monthDict.exists(m) Then
                If dict.exists(m) Then
                    wsReport.Cells(reportRow, col).Value = dict(m)
                Else
                    wsReport.Cells(reportRow, col).Value = 0
                End If
                col = col + 1
            End If
        Next m

        totalColIndex = col
        wsReport.Cells(reportRow, totalColIndex).Value = dict("Total"): col = col + 1
        wsReport.Cells(reportRow, col).Value = dict("FY")

        ' Highlight row if total is negative
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