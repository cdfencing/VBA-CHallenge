Attribute VB_Name = "Module1"
Sub Stocks()

Dim sheet As Worksheet
Dim starting_sheet As Worksheet
Set starting_sheet = ActiveSheet
For Each sheet In ThisWorkbook.Worksheets
    sheet.Activate
'----------------------------------------------------------
'set the variables
Dim ws As Worksheet
Dim i As Long
Dim j As Double
Dim lastrow As Long
Dim Open_Value As Double
Dim Close_Value As Double
Dim yearly_change As Double
Dim Percent_Change As Double
Dim Stock_Volume As Double
Dim Ticker As String
    
    
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly_Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Range("K1").Value = "Percent Change"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"


Stock_Volume = 0
j = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
For i = 2 To lastrow
    Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
        'check to see if ticker name is the same, ie finding the first row of each ticker name
        'then finds the opening value for each ticker name
If (Cells(i, 1).Value <> Cells(i - 1, 1).Value) Then
    Open_Value = Cells(i, 3).Value
    Ticker = Cells(i, 1).Value
    Cells(j, "I") = Ticker
End If
        
        'find the last row for each ticker name
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Close_Value = Cells(i, 6).Value
        
            'calculate the yearly change and put in summary table
        
            yearly_change = (Close_Value - Open_Value)
            Cells(j, 10).Value = yearly_change
        
        
            'checking to see if value is positive or negative
            'if value is negative it is colored red, green if positive
        
            If Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
        
            ElseIf Cells(j, 10).Value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
    
            End If
        
            'calculate the Percent change and add to the first summary table
                If (Close_Value = 0 And Open_Value = 0) Then
                Percent_Change = 0
                
                ElseIf (Close_Value <> 0 And Open_Value = 0) Then
                Percent_Change = 1
                
                Else: Percent_Change = yearly_change / Open_Value
                Cells(j, 11).Value = Percent_Change
                
                End If
            
            'Add the stock volume to the first summary table
        
            Cells(j, 12).Value = Stock_Volume
        
            Stock_Volume = 0
            Percent_Change = 0
            yearly_change = 0
            j = j + 1
    End If
Next i
    
'creating second summary table
Dim lastRow2 As Long
Dim lastRow3 As Long
Dim lastRow4 As Long
        
lastRow2 = Cells(Rows.Count, 11).End(xlUp).Row
        
For Z = 2 To lastRow2
    If Cells(Z, 11) = Application.WorksheetFunction.Max(Range("K2:K" & lastRow2)) Then
    Cells(2, 17).Value = Cells(Z, 11).Value
    Cells(2, 16) = Cells(Z, 9).Value
    
    End If
Next Z

lastRow3 = Cells(Rows.Count, 11).End(xlUp).Row

For y = 2 To lastRow3
   If Cells(y, 11) = Application.WorksheetFunction.Min(Range("K2:K" & lastRow3)) Then
   Cells(3, 17).Value = Cells(y, 11).Value
   Cells(3, 16) = Cells(y, 9).Value
    
    End If
Next y
    
lastRow4 = Cells(Rows.Count, 12).End(xlUp).Row
    
For m = 2 To lastRow4
   If Cells(m, 12) = Application.WorksheetFunction.Max(Range("L2:L" & lastRow4)) Then
   Cells(4, 17).Value = Cells(m, 12).Value
   Cells(4, 16).Value = Cells(m, 9).Value
    
   End If
Next m

               
sheet.Cells(1, 1) = 1 'sets cell A1 to each sheet to 1
Next
starting_sheet.Activate 'Activate the worksheet that was originally active
    
End Sub
