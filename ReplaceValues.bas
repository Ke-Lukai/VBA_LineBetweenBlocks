Attribute VB_Name = "Modul1"
Sub ReplaceValues()

Dim i As Integer
Dim LastRow As Integer

LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To LastRow

    If ActiveSheet.Cells(i, 2) = "A" Then
    
        If ActiveSheet.Cells(i, 3) = "X" Or ActiveSheet.Cells(i, 3) = "Y" Then
        ActiveSheet.Cells(i, 3).Value = "X or Y "
        Else
        ActiveSheet.Cells(i, 3).Value = "Else"
        End If
        
     End If
     
Next i
    
End Sub
