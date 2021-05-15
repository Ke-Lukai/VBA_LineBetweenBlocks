# VBA LineBetweenBlocks
This makro separates blocks of rows that have a different value in RelevantColumn for two adjacent rows. 

Before

![grafik](https://user-images.githubusercontent.com/78645935/118372721-4537f280-b5b3-11eb-9aa7-a5a30c253593.png)

After

![grafik](https://user-images.githubusercontent.com/78645935/118372753-67317500-b5b3-11eb-85e1-9db96fd2517a.png)



Sub multiple_find_and_replace()
    Dim Wbk As Workbook: Set Wbk = ThisWorkbook
    Dim Wrd As Object
    Set Wrd = CreateObject("Word.Application")
    Dim Dict As Object
    Dim RefList As Range
    Dim RefElem As Range
    Wrd.Visible = True
    Dim WDoc As Object
    Set WDoc = Wrd.Documents.Open("C:\Users\Anh\test.docx") 'Your Word document - rename and modify the path.
    Set Dict = CreateObject("Scripting.Dictionary")
    Set RefList = Wbk.Sheets("Sheet1").Range("A1:A4") 'Range of your strings in the database - modify this.

    With Dict
        For Each RefElem In RefList
            If Not .Exists(RefElem) And Not IsEmpty(RefElem) Then
                .Add RefElem.Value, RefElem.Offset(0, 1).Value
            End If
        Next RefElem
    End With

    For Each Key In Dict
        With WDoc.Content.Find
            .Execute FindText:=Key, ReplaceWith:=Dict(Key)
        End With
    Next Key
End Sub
