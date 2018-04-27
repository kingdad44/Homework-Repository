Sub Sheet_Volumes()
  Dim ws As Worksheet
  
  Dim iIndex As Integer
  Dim ts As String
  Dim lastrow
  Dim Volume_Total As Double
  Dim Summary_Table_Row As Integer
 For Each ws In ActiveWorkbook.Worksheets
  Summary_Table_Row = 2
  lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
  For i = 2 To lastrow
    Volume_Total = 0
    ' check ticker against previous value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
      ws.Cells(1, 9).Value = "Ticker Symbol"
      ws.Cells(1, 10).Value = "Annual Volume"
      ' Set the Ticker Symbol
      ts = ws.Cells(i, 1).Value
        
      ' Add to the Volume
      Volume_Total = Volume_Total + ws.Cells(i, 7).Value

      ' Print the Ticker Symbol in the Summary Table
      ws.Range("i" & Summary_Table_Row).Value = ts

      ' Print the Volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Volume_Total


                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total
      Volume_Total = 0

    ' If the cell immediately following a row is the same ts...
    Else

      ' Add to the volume
      Volume_Total = Volume_Total + Cells(i, 3).Value

    End If

  Next i



Next ws


End Sub

