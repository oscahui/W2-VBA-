Sub VBA_Wall_Street()
  Dim ticker_name As String
  Dim ticker_total As Double
  Dim Summary_Table_Row As Integer
  Dim ticker_start As Double
  Dim ticker_end As Double
  Dim yearly_change As Double
  Dim percent_change As Double
  Dim i As Long
  Dim a As Long
  Dim day As Integer
  i = 2
  
  Summary_Table_Row = 2
  LastRow_ticker = Cells(Rows.Count, "A").End(xlUp).Row
  LastRow_challenges = Cells(Rows.Count, "I").End(xlUp).Row
  
  For i = 2 To LastRow_ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ticker_start = Range("C" & i - day)
      ticker_end = Range("F" & i)
      yearly_change = ticker_end - ticker_start
      If ticker_start > 0 And ticker_end <> ticker_start Then
      percent_change = (ticker_end - ticker_start) / ticker_start
      End If
      ticker_name = Cells(i, 1).Value
      ticker_total = ticker_total + Cells(i, 7).Value
      Range("I" & Summary_Table_Row).Value = ticker_name
      If yearly_change < 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        Else
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        End If
      Range("J" & Summary_Table_Row).NumberFormat = "0.00"
      Range("J" & Summary_Table_Row).Value = yearly_change
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      Range("K" & Summary_Table_Row).Value = percent_change
      Range("L" & Summary_Table_Row).Value = ticker_total
      Summary_Table_Row = Summary_Table_Row + 1
      ticker_total = 0
      day = 0
    Else
    ticker_total = ticker_total + Cells(i, 3).Value
    day = day + 1
    End If
  Next i
  
  For a = 2 To LastRow_challenges
   If Cells(a, 11).Value > Max Then
   Max = Cells(a, 11).Value
   max_ticker = Cells(a, 9)
   End If
      
   If Cells(a, 11).Value < Min Then
   Min = Cells(a, 11).Value
   min_ticker = Cells(a, 9)
   End If
   
   If Cells(a, 12).Value > Great_vol Then
   Great_vol = Cells(a, 12).Value
   Great_vol_ticker = Cells(a, 9)
   End If
   
  Next a
  Range("Q" & 2).NumberFormat = "0.00%"
  Range("Q" & 3).NumberFormat = "0.00%"
  Range("P" & 2).Value = max_ticker
  Range("P" & 3).Value = min_ticker
  Range("P" & 4).Value = Great_vol_ticker
  Range("Q" & 2).Value = Max
  Range("Q" & 3).Value = Min
  Range("Q" & 4).Value = Great_vol

End Sub



