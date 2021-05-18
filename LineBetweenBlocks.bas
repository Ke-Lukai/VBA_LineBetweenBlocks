Sub LineBetweenBlocks()

Dim LastRow As Integer
Dim RelevantColumn As Integer

RelevantColumn = 2

LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Integer

    For i = LastRow To 3 Step -1

        If Cells(i, RelevantColumn) <> Cells(i - 1, RelevantColumn) Then
        Rows(i).Insert Shift:=xlDown
        Else
        End If
    
     Next i

End Sub