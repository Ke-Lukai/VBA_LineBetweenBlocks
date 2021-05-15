# VBA LineBetweenBlocks
This makro separates blocks of rows that have a different value in RelevantColumn for two adjacent rows. 

Before

![grafik](https://user-images.githubusercontent.com/78645935/118372721-4537f280-b5b3-11eb-9aa7-a5a30c253593.png)

After

![grafik](https://user-images.githubusercontent.com/78645935/118372753-67317500-b5b3-11eb-85e1-9db96fd2517a.png)



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
