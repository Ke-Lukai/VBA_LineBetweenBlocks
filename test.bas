Option Explicit
Dim Source As Workbook
Dim Destination As Workbook

Sub CopyPasteFilteredLines()

'Set requires an opened workbook

'-------------------------------------------------------------
'Trick, dass es auf der gleichen Ebene wie die anderen Dateien sein muss.
'-------------------------------------------------------------

Set Source = Workbooks.Open("C:\VBA\GitHub Portfolio\CopyPasteFilteredLines\Source.xlsx")
    
'Set requires an opened workbook
Set Destination = Workbooks.Open("C:\VBA\GitHub Portfolio\CopyPasteFilteredLines\Destination.xlsx")
            
            
ClearDestination

Source.Worksheets(1).Cells(1, 1).CurrentRegion.AutoFilter Field:=2, Criteria1:=Array("A")
Source.Worksheets(1).Cells(1, 1).CurrentRegion.AutoFilter Field:=1, Criteria1:=Array("1", "3", "5", "7", "9")



'-------------------------------------------------------------
'Findet Column Index für gegebenen ColumnHead, um später bei Filter nicht Index, sondern ColumnHead verwenden zu können
'-------------------------------------------------------------
'Sub test()
' Dim r As Long
' r = Rows("1").Find("Current State").Column ' use whichever row contains headings, I assumed Row 1
'
' ' then when you get to the point in your macro that sets autofilter use r as column number
'End Sub
'-------------------------------------------------------------

Source.Worksheets(1).Cells(1, 1).CurrentRegion.Copy Destination:=Destination.Worksheets(1).Cells(1, 1)

'-----------------------------------------------------------------------
'Liest eindeutige Liste aus Spalte ein
'-----------------------------------------------------------------------

'Sub UniqueVals()
Dim i As Variant
Dim j As Variant

j = Application.Transpose(Range("A1", Range("A" & Rows.Count).End(xlUp)))

With CreateObject("Scripting.Dictionary")
For Each i In j
.Item(i) = i
Next
Cells(1, 2).Resize(.Count) = Application.Transpose(.Keys)
End With
End Sub










End Sub