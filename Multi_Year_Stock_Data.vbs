Attribute VB_Name = "Module1"
Sub MultiYearStock_Data()
  
  For Each Ws In Worksheets
  
  Dim WorkSheetName As String
  WorkSheetName = Ws.Name
  
  Dim Ticker As String
  Dim TotalStockVolumn As Double
  TotalStockVolumn = 0
  
  LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  Dim SummaryTableRow As Integer
  Dim YearOpen As Double
  Dim YearClose As Double
  
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "YearlyChange"
  Cells(1, 11).Value = "PercentageChange"
  Cells(1, 12).Value = "TotalStockVolumn"
  
  SummaryTableRow = 2
  
  'Loop through all tickers in Column A
  For i = 2 To LastRow
  
    Ticker = Cells(i, 1).Value
  
   If YearOpen = 0 Then
   
     YearOpen = Cells(i, 3).Value
     
   End If
   
   'Check if we are still within the same ticker
   If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
     
     YearClose = Cells(i, 6).Value
     
     'Yearly Change
     YearlyChange = YearClose - YearOpen
     
     'Percent Change
     If YearOpen = 0 Then
        PercentageChange = 0
     Else
        PercentageChange = YearlyChange / YearOpen
     End If
     
       TotalStockVolumn = TotalStockVolumn + Cells(i, 7).Value
       
     
       Ws.Range("I" & SummaryTableRow).Value = Ticker
     
       Ws.Range("J" & SummaryTableRow).Value = YearlyChange
     
       Ws.Range("K" & SummaryTableRow).Value = PercentageChange
     
       Ws.Range("L" & SummaryTableRow).Value = TotalStockVolumn
       
       'Conditional Formatting
       If Cells(SummaryTableRow, 10).Value > 0 Then
          Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
       ElseIf Cells(SummaryTableRow, 10).Value < 0 Then
          Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
       End If
       
        SummaryTableRow = SummaryTableRow + 1
     
        TotalStockVolumn = 0
     
    End If
    
  Next i
  
 Next Ws
     
End Sub
