
Sub DistinctCollection()

Dim coll As Collection
Set coll = New Collection

'Marks column, whose values get collected. Column A would be RelevantColumnIndex = 1
Dim RelevantColumnIndex As Long
Dim FirstRowData As Long

'Setting variables
RelevantColumnIndex = 2
FirstRowData = 2

Dim LastRow As Integer
LastRow = ActiveSheet.Cells(Rows.Count, RelevantColumnIndex).End(xlUp).Row

Dim i As Integer
'If entry is not already in collection, it gets added to collection.
For i = FirstRowData To LastRow
    If AlreadyInCollection(coll, i, RelevantColumnIndex) = False Then
    coll.Add ActiveSheet.Cells(i, RelevantColumnIndex).Value
    Else
    End If
    
Next i

End Sub



'Tests, whether entry is already in collection
Function AlreadyInCollection(coll, i, RelevantColumnIndex) As Boolean
Dim x As Integer
        For x = 1 To coll.Count
                If coll(x) = ActiveSheet.Cells(i, RelevantColumnIndex).Value Then
                AlreadyInCollection = True
                Exit For
                Else
                AlreadyInCollection = False
                End If
        Next x
End Function



