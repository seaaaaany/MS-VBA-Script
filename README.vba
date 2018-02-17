Sub Easy():

    Dim WS_Count As Integer
    Dim j As Integer
    Dim starting_ws As Worksheets
    Set starting_ws=ActiveSheet

    WS_Count = ActiveWorkbook.Worksheets.Count

    For j = 1 To WS_Count

    ThisWorkbook.Worksheets(j).Activate

    Dim x As Double
    Dim Total As Double
    Dim TotalV As Double

        Columns("I:Q").Select
        Selection.Clear

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Value"

        x = 2
        Cells(x, 9).Value = Cells(x, 1).Value

        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        For I = 2 To LastRow

        If Cells(I, 1).Value = Cells(x, 9).Value Then

            TotalV = TotalV + Cells(I, 7).Value

        Else

            Cells(x, 10).Value = TotalV
            TotalV = Cells(I, 7).Value
            x = x + 1
            Cells(x, 9).Value = Cells(I, 1).Value

        End If

    Next I

    Cells(x, 10).Value = TotalV


    Columns("I:Q").EntireColumn.AutoFit
    Cells(1, 1).Select

    Next j


End Sub
