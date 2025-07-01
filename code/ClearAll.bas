Sub ClearAll()

    Sheets("Revenue Report").Cells.Clear

    Sheets("Data").Cells.Clear

   

    ' Clear previous data from TR Template, A3 down and R2, S2

    Set wsDest = Sheets("TR Template")

    wsDest.Range("A3:U" & wsDest.Rows.Count).ClearContents

    wsDest.Range("R2").ClearContents

    wsDest.Range("S2").ClearContents

    MsgBox "'Data', 'Revenue Report' and 'TR Template' Sheet has been cleared . Please add Data in the 'Data' Sheet to procceed.", vbInformation

End Sub