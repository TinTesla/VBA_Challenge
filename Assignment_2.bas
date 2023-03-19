Attribute VB_Name = "Assignment_2"
Sub Assignmet_2()

' Placeholders
Dim Ticker_Name As String

Dim Open_Total As Double
Open_Total = 0
Dim Close_Total As Double
Close_Total = 0
Dim Volume_Total As Double
Volume_Total = 0
Dim P_Change As Double

Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
  
    'Column Labels + Formatting
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 9).Columns.AutoFit
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 10).Columns.AutoFit
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 11).Columns.AutoFit
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 12).Columns.AutoFit
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 15).Columns.AutoFit
    ws.Cells(1, 16).Value = "Results"
    ws.Cells(1, 16).Columns.AutoFit
    ws.Cells(2, 14).Value = "Greatest Percent Increase"
    ws.Cells(2, 14).Columns.AutoFit
    ws.Cells(3, 14).Value = "Greatest Percent Decrease"
    ws.Cells(3, 14).Columns.AutoFit
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(4, 14).Columns.AutoFit
    'Starting Value for Bonus
    ws.Cells(3, 16).Value = "1000%"
    
    'Row Check
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Row Memory For Data_Dump
    Dim Data_Dump As Integer
    Data_Dump = 2

    'Data Crunch
    For i = 2 To Last_Row

        ' Data Dump Loop
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

            'Values Added
            Ticker_Name = ws.Cells(i, 1).Value
            Open_Total = Open_Total + ws.Cells(i, 3).Value
            Close_Total = Close_Total + ws.Cells(i, 6).Value
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value

            ' Print Combined Results
            ws.Range("I" & Data_Dump).Value = Ticker_Name
            ws.Range("J" & Data_Dump).Value = Open_Total - Close_Total
            P_Change = Open_Total / Close_Total
            P_Change = P_Change - 1
            ws.Range("K" & Data_Dump).Value = P_Change
                'Greatest Increase Check
                If P_Change >= ws.Range("P2").Value Then
                    ws.Range("O2").Value = Ticker_Name
                    ws.Range("P2").Value = P_Change
                'Greatest Decrease Check
                ElseIf P_Change < ws.Range("P3").Value Then
                    ws.Range("O3").Value = Ticker_Name
                    ws.Range("P3").Value = P_Change
                End If
            ws.Range("L" & Data_Dump).Value = Volume_Total
                'Greatest Volume Check
                If Volume_Total > ws.Range("P4").Value Then
                    ws.Range("O4").Value = Ticker_Name
                    ws.Range("P4").Value = Volume_Total
                End If
                
            'Color Coding % Change
            If ws.Range("J" & Data_Dump).Value >= 0 Then
            
                'Change to Green
                ws.Range("J" & Data_Dump).Interior.ColorIndex = 4
                
            ElseIf ws.Range("J" & Data_Dump).Value < 0 Then
                
                'Change to Red
                ws.Range("J" & Data_Dump).Interior.ColorIndex = 3
                
            End If
            
            ' Data_Dump update
            Data_Dump = Data_Dump + 1
      
            ' Placeholders Reset
            Open_Total = 0
            Close_Total = 0
            Volume_Total = 0
        
        'Value Added Loop
        Else

            'Values Added
            Open_Total = Open_Total + ws.Cells(i, 3).Value
            Close_Total = Close_Total + ws.Cells(i, 6).Value
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value

        End If

    Next i
    
Next ws

'Bonus

'Formatting
Sheets("2018").Range("N1:P4").Copy Destination:=Sheets("2018").Range("N6")
Sheets("2018").Range("N6").Value = "_2018_:"
Sheets("2019").Range("N1:P4").Copy Destination:=Sheets("2018").Range("N11")
Sheets("2018").Range("N11").Value = "_2019_:"
Sheets("2020").Range("N1:P4").Copy Destination:=Sheets("2018").Range("N16")
Sheets("2018").Range("N16").Value = "_2020_:"
Sheets("2018").Range("N1").Value = "All Sheets:"

For T = 6 To 29

    'Greatest % Increase All Sheets
    If Sheets("2018").Cells(T, 14).Value = "Greatest % Increase" Then
        If Sheets("2018").Cells(2, 16).Value < Sheets("2018").Cells(T, 16).Value Then
          Sheets("2018").Cells(2, 15).Value = Sheets("2018").Cells(T, 15).Value
          Sheets("2018").Cells(2, 16).Value = Sheets("2018").Cells(T, 16).Value
        End If
    End If
                  
    'Greatest % Decrease All Sheets
    If Sheets("2018").Cells(T, 14).Value = "Greatest % Decrease" Then
       If Sheets("2018").Cells(3, 16).Value > Sheets("2018").Cells(T, 16).Value Then
           Sheets("2018").Cells(3, 15).Value = Sheets("2018").Cells(T, 15).Value
           Sheets("2018").Cells(3, 16).Value = Sheets("2018").Cells(T, 16).Value
        End If
    End If

    'Greatest Total Volume Per Sheet
    If Sheets("2018").Cells(T, 14).Value = "Greatest Total Volume" Then
        If Sheets("2018").Cells(4, 16).Value < Sheets("2018").Cells(T, 16).Value Then
          Sheets("2018").Cells(4, 15).Value = Sheets("2018").Cells(T, 15).Value
          Sheets("2018").Cells(4, 16).Value = Sheets("2018").Cells(T, 16).Value
        End If
    End If
    
Next T

End Sub


