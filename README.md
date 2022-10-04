# Stock_analysis
Below are screenshots of the partial results for each worksheet in the Multiple_year_stock_data file, as well as a screenshot of the VBA script. The VBA script has been uploaded to the repository as a "BAS" file. 
## Results for 2018
![image](results_screenshots/2018.png)
<br>
## Results for 2019
![image](results_screenshots/2019.png)
<br>
## Results for 2020
![image](results_screenshots/2020.png)
<br>
## VBA Script

Sub loop_all_worksheets()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        stock_analysis
    Next ws
End Sub

Sub stock_analysis()

    'omar espinoza
  
    Dim ticker As String
    Dim opening As Double
    Dim closing As Double
    Dim volume As Double
    Dim i As Long
    Dim j As Long
    Dim lrow As Long
    Dim maxincrease As Double
    Dim maxdecrease As Double
    Dim maxvolume As Double
    Dim maxincreaseticker As String
    Dim maxidecreaseticker As String
    Dim maxvolumeticker As String
    
    'add headers and titles and adjusts column width
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total stock volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    j = 2 'row counter for the output table
    maxdecrease = 0
    maxincrease = 0
    maxvolume = 0
    
    'initializes values with first row of data
    ticker = Cells(2, 1).Value
    volume = Cells(2, 7).Value
    opening = Cells(2, 3).Value
     
    For i = 2 To lrow
        If (Cells(i + 1, 1).Value = ticker) Then 'checks if next row corresponds with current ticker
            volume = volume + Cells(i + 1, 7).Value
            If volume > maxvolume Then 'bonus
                maxvolume = volume
                maxvolumeticker = ticker
            End If
        Else 'if next row is different ticker, it posts values and start new ticker
            closing = Cells(i, 6).Value
            Cells(j, 12).Value = volume
            Cells(j, 9).Value = ticker
            Cells(j, 10).Value = closing - opening
                If (Cells(j, 10).Value < 0) Then 'conditional formatting
                    Cells(j, 10).Interior.ColorIndex = 3
                ElseIf (Cells(j, 10).Value > 0) Then
                        Cells(j, 10).Interior.ColorIndex = 4
                Else: Cells(j, 10).Interior.ColorIndex = 0
                End If
            Cells(j, 11).Value = FormatPercent((closing - opening) / opening, 2)
            If Cells(j, 11).Value < 0 Then 'checks if %decrease is lower thatn current greatest decrease
                If Cells(j, 11).Value < maxdecrease Then
                    maxdecrease = Cells(j, 11).Value
                    maxdecreaseticker = ticker
                End If
            ElseIf Cells(j, 11).Value > 0 Then 'checks if %decrease is lower thatn current greatest increase
                If Cells(j, 11).Value > maxincrease Then
                    maxincrease = Cells(j, 11).Value
                    maxincreaseticker = ticker
                End If
            End If
            volume = Cells(i + 1, 7)
            ticker = Cells(i + 1, 1).Value
            opening = Cells(i + 1, 3).Value
            j = j + 1
        End If
    Next i
    'posts greatest increase, decrease, and volume info
    Cells(2, 15).Value = maxincreaseticker
    Cells(3, 15).Value = maxdecreaseticker
    Cells(4, 15).Value = maxvolumeticker
    Cells(2, 16).Value = FormatPercent(maxincrease, 2)
    Cells(3, 16).Value = FormatPercent(maxdecrease, 2)
    Cells(4, 16).Value = maxvolume
    
    'Autoadjusts all columns' width
    Columns("i:l").AutoFit
    Columns("n:p").AutoFit
        
End Sub