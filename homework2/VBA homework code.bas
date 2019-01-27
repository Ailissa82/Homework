Sub MultipleYearStock()

' delcare variables

Dim LastRowNum, FirstTickerRowNum, NextTickerRowNum As Long
Dim TickerYearClose, TickerYearOpen As Double
Dim TickerNum As Long
Dim TickerName, NextTickerName  As String
Dim Volume As Double
Dim TotalTabNum As Integer

' count the number of worksheets in the workbook
TotalTabNum = ThisWorkbook.Sheets.Count

'create loop on worksheets, activeate worksheet, set headers, count rows, set base value

For j = 1 To TotalTabNum

  Sheets(j).Activate
  TickerNum = 1
  FirstTickerRowNum = 2
  Range("M1").Value = "Ticker"
  Range("N1").Value = "Yearly Change"
  Range("O1").Value = "Percent Change"
  Range("P1").Value = "Total Stock Volume"
  LastRowNum = Cells(Rows.Count, 1).End(xlUp).Row
  Volume = Cells(2, 7).Value
  
'customize within worksheet
  ' start at row 2 and go through the last number on each tab
  For i = 2 To LastRowNum

    TickerName = Cells(i, 1).Value
    NextTickerName = Cells(i + 1, 1).Value
    ' add the volume of all like ticker names
    If TickerName = NextTickerName Then
      Volume = Volume + Cells(i + 1, 7).Value
    Else
       TickerNum = TickerNum + 1
       NextTickerRowNum = i + 1
       ' putting ticker and volume in the summary
       Cells(TickerNum, 13).Value = TickerName
       Cells(TickerNum, 16).Value = Volume
       ' pulling closing and opening value per ticker
       TickerYearClose = Cells(NextTickerRowNum - 1, 6).Value
       TickerYearOpen = Cells(FirstTickerRowNum, 3).Value
       
       ' looking for where the opening value is 0
       While TickerYearOpen = 0
             FirstTickerRowNum = FirstTickerRowNum + 1
             TickerYearOpen = Cells(FirstTickerRowNum, 3).Value
       Wend
       ' calculate yaearly change
       Cells(TickerNum, 14).Value = TickerYearClose - TickerYearOpen
       Cells(TickerNum, 15).Value = Cells(TickerNum, 14).Value / TickerYearOpen
       Cells(TickerNum, 15).NumberFormat = "0.00%"
       
        ' adding conditional formatting
       If Cells(TickerNum, 14).Value > 0 Then
            Cells(TickerNum, 14).Interior.ColorIndex = 4
        Else
            Cells(TickerNum, 14).Interior.ColorIndex = 3
        End If
        
    ' move to next row to get above data for the next ticker and resetting the start vale of volume
      FirstTickerRowNum = NextTickerRowNum
      Volume = Cells(FirstTickerRowNum, 7).Value
    End If
' moves to the next ticker
 Next i

 Range("M1:P1").Columns.AutoFit
' move to next tab on the worksheet
Next j

End Sub





