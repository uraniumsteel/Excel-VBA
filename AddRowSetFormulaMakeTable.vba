
Sub AddRowSetFormulaMakeTable()
'
' Select then rename the sheet called "Projects by Status" to DataList
    Sheets("Projects by Status - Data List").Select
    Sheets("RTO Projects by Sta - Data List").Name = "DataList"

    'Activate that sheet and select what we know is the an empty column
    Sheets("DataList").Activate
    Range("O1").Select                 ' Empty column at end of data we have
    ActiveCell.FormulaR1C1 = "Days Since Last Updated"  'Set title for that column

    'Get Last Cell in Row, then put a formula into every cell in that new row we
    '  created above
    Range("N2").Select  'Select a row with all data - no blank rows
    lRow = Selection.End(xlDown).Row 'Go to the end and get the row #
    Range("O2:O" & lRow).Formula = "=Today()-K2"  'Put forumla into every cell in that column

    'Select the whole set of data and make a table called "Table1"
    Range("A1:O" & lRow).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:O" & lRow), , xlYes).Name = _
        "Table1"
    Range("Table1[#All]").Select   'Select the table.

End Sub
