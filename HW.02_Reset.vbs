Sub Reset()
For Each ws In Worksheets
ws.Activate

Columns("I:P").Value = ""
Columns("I:P").Interior.ColorIndex = 0

Next ws

End Sub

