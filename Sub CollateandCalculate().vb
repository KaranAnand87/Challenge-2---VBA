Sub CollateandCalculate()

'Combining Data
    Sheets.Add.Name = "Collated Stock Data"
    Sheets("Collated Stock Data").Move Before:=Sheets(1)
    Set combined_sheet = Worksheets("Collated Stock Data")

    For Each ws In Worksheets

        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        lastRowYear = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

        combined_sheet.Range("A" & lastRow & ":G" & ((lastRowYear - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowYear + 1)).Value

    Next ws

    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    combined_sheet.Range("J1").Value = "Ticker"
    combined_sheet.Range("K1").Value = "Yearly Change"
    combined_sheet.Range("L1").Value = "Percent Change"
    combined_sheet.Range("M1").Value = "Total Stock Volume"
    combined_sheet.Range("Q1").Value = "Ticker"
    combined_sheet.Range("R1").Value = "Value"
    combined_sheet.Range("P2").Value = "Greatest % Increase"
    combined_sheet.Range("P3").Value = "Greatest % Decrease"
    combined_sheet.Range("P4").Value = "Greatest Total Volume"
    
     combined_sheet.Columns("A:R").AutoFit

'Defining variables
  Dim Ticker As String
  Dim OpenPrice As Double
  Dim ClosePrice As Double
  Dim YearlyChange As Double
  Dim PercentChange As Double
  Dim TotalStockVolume As Double
  Dim SummaryTableRow As Double
  
  'Code for Retrieval of Data and Column Creation
  
  SummaryTableRow = 2
  TickerNumber = 1
  TotalStockVolume = 0

  lastRow = Cells(Rows.Count, 1).End(xlUp).Row

  For i = 2 To lastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    
    TickerNumber = TickerNumber + 1
    
    OpenPrice = Cells(TickerNumber, 3).Value
    ClosePrice = Cells(i, 6).Value
    
    YearlyChange = ClosePrice - OpenPrice
    
    PercentChange = YearlyChange / OpenPrice
    

    Range("J" & SummaryTableRow).Value = Ticker
    Range("K" & SummaryTableRow).Value = YearlyChange
    Range("L" & SummaryTableRow).Value = PercentChange
    Range("L" & SummaryTableRow).NumberFormat = "0.00%"
    Range("M" & SummaryTableRow).Value = TotalStockVolume
    
    SummaryTableRow = SummaryTableRow + 1
    
    TotalStockVolume = 0
    YearlyChange = 0
    PercentChange = 0
    TickerNumber = i
    
    Else
    
    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

    End If

  Next i

'Code for Conditional Formatting

LastKRow = Cells(Rows.Count, 11).End(xlUp).Row

For j = 2 To LastKRow

If Cells(j, 11).Value > 0 Then

Cells(j, 11).Interior.ColorIndex = 4

Else

Cells(j, 11).Interior.ColorIndex = 3

End If

Next j

LastLRow = Cells(Rows.Count, 12).End(xlUp).Row

For k = 2 To LastLRow

If Cells(k, 12).Value > 0 Then

Cells(k, 12).Interior.ColorIndex = 4

Else

Cells(k, 12).Interior.ColorIndex = 3

End If

Next k

'Code for Calculated Values

Increase = 0
Decrease = 0
Greatest = 0

For l = 3 To LastLRow

Last_l = l - 1

Current_l = Cells(l, 12).Value

Previous_l = Cells(Last_l, 12).Value

Volume = Cells(l, 13).Value

PreviousVolume = Cells(Last_l, 13).Value

If Increase > Current_l And Increase > Previous_l Then

Increase = Increase

ElseIf Current_l > Increase And Current_l > Previous_l Then
Increase = Current_l
Range("R2").Value = Increase
Range("R2").NumberFormat = "0.00%"
Range("Q2").Value = Cells(l, 10).Value

ElseIf Previous_l > Increase And Previous_l > Current_l Then
Increase = Previous_l
Range("R2").Value = Increase
Range("R2").NumberFormat = "0.00%"
Range("Q2").Value = Cells(l, 10).Value

End If

If Decrease < Current_l And Decrease < Previous_l Then

Decrease = Decrease

ElseIf Current_l < Increase And Current_l < Previous_l Then
Decrease = Current_l
Range("R3").Value = Decrease
Range("R2").NumberFormat = "0.00%"
Range("Q3").Value = Cells(l, 10).Value

ElseIf Previous_l < Decrease And Previous_l < Current_l Then
Decrease = Previous_l
Range("R3").Value = Increase
Range("R3").NumberFormat = "0.00%"
Range("Q3").Value = Cells(l, 10).Value

End If



If Greatest > Volume And Greatest > PreviousVolume Then

Greatest = Greatest

ElseIf Volume > Greatest And Volume > PreviousVolume Then
Greatest = Volume
Range("R4").Value = Greatest
Range("Q4").Value = Cells(l, 10).Value

ElseIf PreviousVolume > Greatest And PreviousVolume > Volume Then
Greatest = PreviousVolume
Range("R4").Value = Greatest
Range("Q4").Value = Cells(l, 10).Value

End If

Next l

End Sub




