Attribute VB_Name = "Module1"
Sub Stock()

Dim ws As Worksheet
' Set ws = ActiveSheet
For Each ws In ActiveWorkbook.Sheets
    Dim RowNumber As Long
    Dim CurrentTicker As String
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim CurrentOpenValue As Double
    Dim CurrentCloseValue As Double
    Dim IsPastFirstLine As Boolean
    LastRow = Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim SummaryRN As Integer
    SummaryRN = 2
  
  ' Adding this feature to accommodate the calculations to appear in each tab
    CurrentTicker = ""
    TotalVolume = 0
    CurrentOpenValue = 0
    CurrentCloseValue = 0
    IsPastFirstLine = False
    
 'Deleting the range prior to runing the code to ensure clean data fill
 
    ws.Range("I:Q").Delete
    
 'Adding titles to the new columns
 
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percentage Change"
            ws.Range("L1").Value = "Total Stock Volume"
                     
  'Applying Autofit
            
    ws.Range("I:Q").Columns.AutoFit
        
    For RowNumber = 2 To (LastRow + 1)
        Dim TickerCell As Range
        Dim VolumeCell As Range
        Dim OpenCell As Range
        Dim CloseCell As Range
        Set TickerCell = ws.Cells(RowNumber, "A")
        Set VolumeCell = ws.Cells(RowNumber, "G")
        Set OpenCell = ws.Cells(RowNumber, "C")
        
        If CurrentTicker <> TickerCell.Value Then
            
            If IsPastFirstLine = True Then
                Set CloseCell = ws.Cells(RowNumber - 1, "F")
                CurrentCloseValue = CloseCell.Value
                Dim YearlyChangeValue As Double
                YearlyChangeValue = CurrentCloseValue - CurrentOpenValue
                Dim PercentageChangeValue As Double
                PercentageChangeValue = YearlyChangeValue / CurrentOpenValue
                
              
            ws.Cells(SummaryRN, "I") = CurrentTicker
            ws.Cells(SummaryRN, "J") = YearlyChangeValue
            ws.Cells(SummaryRN, "K") = PercentageChangeValue
            ws.Cells(SummaryRN, "L") = TotalVolume
            SummaryRN = SummaryRN + 1
             
            End If
                
            CurrentOpenValue = OpenCell.Value
            CurrentTicker = TickerCell.Value
            TotalVolume = 0
                  
         End If
         
         IsPastFirstLine = True
         TotalVolume = TotalVolume + VolumeCell.Value
    Next
    
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("J:J").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    ws.Range("J:J").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    ws.Range("J:J").FormatConditions(1).Interior.Color = vbGreen
    ws.Range("J:J").FormatConditions(2).Interior.Color = vbRed
    ws.Range("J1").FormatConditions.Delete
    
Next
End Sub



