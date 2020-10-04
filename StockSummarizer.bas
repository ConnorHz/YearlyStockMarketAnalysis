Attribute VB_Name = "StockSummarizer"

Sub Main()
    
    
    ' Disable Screen Updating. Gotta Go Fast
    Application.ScreenUpdating = False

    Dim daily As New StockDaily
    Dim summary As New StockSummary
    Dim dailys As New Collection
    Dim summaries As New Collection
    
    Dim nextTicker As String
    Dim lastrow As Long
    Dim y As Long
    
    Dim firstDaily As New StockDaily
    Dim lastDaily As New StockDaily
    Dim TotalVolume As Double
    
    Dim foo As String
    
    For Each ws In ActiveWorkbook.Worksheets
    
        ws.Columns("J:M").EntireColumn.Clear
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For y = 2 To lastrow
            
            ' Reset the Daily
            Set daily = New StockDaily
          
            ' Assign Daily Values
            daily.ticker = ws.Range("A" & y).value
            daily.Day = StringToDate(ws.Range("B" & y).value)
            daily.OpenPrice = ws.Range("C" & y).value
            daily.HighPrice = ws.Range("D" & y).value
            daily.LowPrice = ws.Range("E" & y).value
            daily.ClosePrice = ws.Range("F" & y).value
            daily.Volume = ws.Range("G" & y).value
            
            dailys.Add Item:=daily
            
            nextTicker = ws.Range("A" & y + 1).value
                
                ' If it's the same ticker symbol then add it to the dailys, else summarize the ticker and move to the next
                If (daily.ticker <> nextTicker) Then
                
                    ' Calculate and write to summary
                    Dim d As StockDaily
                    For Each d In dailys
                    
                       If (firstDaily.ticker = "" Or firstDaily.Day > d.Day) Then
                            Set firstDaily = d
                       End If
                       
                       If (lastDaily.ticker = "" Or lastDaily.Day < d.Day) Then
                            Set lastDaily = d
                       End If
                       
                       TotalVolume = TotalVolume + d.Volume
                       
                    Next d
                    
                    ' Clear Dailys to save memory
                    Set dailys = New Collection
                    
                    ' Fill out the summary
                    summary.Year = Year(firstDaily.Day)
                    summary.ticker = firstDaily.ticker
                    summary.PriceChange = firstDaily.GetPriceChange(lastDaily)
                    summary.PercentChange = firstDaily.GetPercentChange(lastDaily)
                    summary.TotalVolume = TotalVolume
                    
                    summaries.Add summary
                    
                    ' Reset Variables and Objects
                    TotalVolume = 0
                    Set summary = New StockSummary
                    Set firstDaily = New StockDaily
                    Set lastDaily = New StockDaily
                    
            End If
               
        Next y
           
        Call WriteCollectionToSheet(summaries, ws)
        
        Dim greatestPecentIncrease As StockSummary
        Dim greatestPecentDecrease As StockSummary
        Dim greatestTotalVolume As StockSummary
        Dim firstRun As Boolean
        
        firstRun = True
        
        
        ' Use summary from above. Dangerous, I know, but I'll reset it when I'm done. I'm just borrowing
        Set summary = New StockSummary
        
        For Each summary In summaries
           
            ' If this is the first run though the loop, just set all the "greatest" summaries to the first item
            If (firstRun) Then
            
                Set greatestPercentIncrease = summary
                Set greatestPercentDecrease = summary
                Set greatestTotalVolume = summary
                
                firstRun = False
            Else
                If (summary.PercentChange > greatestPercentIncrease.PercentChange) Then
                    Set greatestPercentIncrease = summary
                End If
                
                If (summary.PercentChange < greatestPercentDecrease.PercentChange) Then
                    Set greatestPercentDecrease = summary
                End If
                
                If (summary.TotalVolume > greatestTotalVolume.TotalVolume) Then
                    Set greatestTotalVolume = summary
                End If
                
            End If
            
        Next summary
        
        ws.Range("O1").value = "Greatest Percent Increase"
        ws.Range("O2").value = greatestPercentIncrease.ticker
        
        ws.Range("P1").value = "Greatest Percent Decrease"
        ws.Range("P2").value = greatestPercentDecrease.ticker
        
        ws.Range("Q1").value = "Greatest Total Volume"
        ws.Range("Q2").value = greatestTotalVolume.ticker
        
        ws.Columns("O:Q").EntireColumn.AutoFit
        
        Set summary = New StockSummary
        
        
        
        ' Clear Summaries to save memory
        Set summaries = New Collection
        
    Next ws

    
    MsgBox ("Done")
    

End Sub

' Converts yyyymmdd formatted String into Date
Function StringToDate(stringDate As String) As Date
    Dim Year As Integer
    Dim month As Integer
    Dim Day As Integer
    
    Year = CInt(Left(stringDate, 4))
    month = CInt(Mid(stringDate, 5, 2))
    Day = CInt(Right(stringDate, 2))
    
    StringToDate = DateSerial(Year:=Year, month:=month, Day:=Day)
    
End Function

Sub WriteCollectionToSheet(table As Collection, ws As Variant)

    Dim dictYear As New Dictionary
    Dim dictTicker As New Dictionary
    Dim dictPriceChange As New Dictionary
    Dim dictPercentChange As New Dictionary
    Dim dictTotalVolume As New Dictionary
    Dim key As Variant
    Dim summaryRange As Range

    Dim i As Long
    For i = 1 To table.Count
        dictYear.Add i, table(i).Year
        dictTicker.Add i, table(i).ticker
        dictPriceChange.Add i, table(i).PriceChange
        dictPercentChange.Add i, FormatPercent(table(i).PercentChange, 0)
        dictTotalVolume.Add i, table(i).TotalVolume
                
    Next i
    
    ws.Range("J1").value = "Ticker Symbol"
    ws.Range("K1").value = "Price Change"
    ws.Range("L1").value = "Percent Change"
    ws.Range("M1").value = "Total Volume"

    ws.Range("J2:J" & UBound(dictTicker.Items) + 2) = WorksheetFunction.Transpose(dictTicker.Items)
    ws.Range("K2:K" & UBound(dictPriceChange.Items) + 2) = WorksheetFunction.Transpose(dictPriceChange.Items)
    ws.Range("L2:L" & UBound(dictPercentChange.Items) + 2) = WorksheetFunction.Transpose(dictPercentChange.Items)
    ws.Range("M2:M" & UBound(dictTotalVolume.Items) + 2) = WorksheetFunction.Transpose(dictTotalVolume.Items)
    
    Set summaryRange = ws.Range("J2:M" & UBound(dictTicker.Items) + 2)
    
    'Format Cells
    ws.Cells.FormatConditions.Delete
    
    ' Set Green if Positive
    summaryRange.FormatConditions.Add Type:=xlExpression, Formula1:="=$K2>0"
    summaryRange.FormatConditions(summaryRange.FormatConditions.Count).SetFirstPriority
    With summaryRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    
    ' Set Red if Negative
    summaryRange.FormatConditions.Add Type:=xlExpression, Formula1:="=$K2<0"
    summaryRange.FormatConditions(summaryRange.FormatConditions.Count).SetFirstPriority
    With summaryRange.FormatConditions(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 7961087
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ws.Columns("K").NumberFormat = "$#,##0.00_);($#,##0.00)"
    ws.Columns("M").NumberFormat = "#,##0"
    ws.Columns("J:M").EntireColumn.AutoFit

End Sub

