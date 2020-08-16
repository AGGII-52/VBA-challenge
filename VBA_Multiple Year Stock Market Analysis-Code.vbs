Sub Stock_Analysis()

Dim Label As String
Dim Printvalue As Double
Dim a As Integer
Dim TotalVol As Double
Dim OpenValue As Double
Dim CloseValue As Double
Dim YrChange As Double
Dim PChange As Double

a = Application.Worksheets.Count

For I = 1 To a

Worksheets(1).Activate
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

Next

Printvalue = 2

    For I = 2 To 797711

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        Label = Cells(I, 1).Value

        TotalVol = TotalVol + Cells(I, 7).Value

        Range("I" & Printvalue).Value = Label
        
        Range("K" & Printvalue).NumberFormat = "0.00"

        Range("L" & Printvalue).Value = TotalVol

        Printvalue = Printvalue + 1

        TotalVol = 0

    Else

        TotalVol = TotalVol + Cells(I, 7).Value

    End If
    
Next
    
OpenValue = Cells(2, 3).Value
Printvalue = 2

    For I = 2 To 797711

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        CloseValue = Cells(I, 6).Value

        YrChange = CloseValue - OpenValue

        Range("J" & Printvalue).Value = YrChange
        
            If YrChange > 0 Then
            
            Range("J" & Printvalue).Interior.ColorIndex = 4
            
            ElseIf YrChange < 0 Then

            Range("J" & Printvalue).Interior.ColorIndex = 3
            
            End If

        PChange = 100 * (CloseValue - OpenValue) / OpenValue
        
        Range("K" & Printvalue).Value = PChange

        Printvalue = Printvalue + 1

        YrChange = 0
    
    End If
    
Next
        
For I = 1 To a

Worksheets(2).Activate
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

Next

Printvalue = 2

    For I = 2 To 760192

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        Label = Cells(I, 1).Value

        TotalVol = TotalVol + Cells(I, 7).Value

        Range("I" & Printvalue).Value = Label
        
        Range("K" & Printvalue).NumberFormat = "0.00"

        Range("L" & Printvalue).Value = TotalVol

        Printvalue = Printvalue + 1

        TotalVol = 0

    Else

        TotalVol = TotalVol + Cells(I, 7).Value

        End If
    
Next
    
OpenValue = Cells(2, 3).Value
Printvalue = 2

    For I = 2 To 760192

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        CloseValue = Cells(I, 6).Value

        YrChange = CloseValue - OpenValue

        Range("J" & Printvalue).Value = YrChange
        
            If YrChange > 0 Then
            
            Range("J" & Printvalue).Interior.ColorIndex = 4
            
            ElseIf YrChange < 0 Then

            Range("J" & Printvalue).Interior.ColorIndex = 3
            
            End If

        PChange = 100 * (CloseValue - OpenValue) / OpenValue
        
        Range("K" & Printvalue).Value = PChange

        Printvalue = Printvalue + 1

        YrChange = 0
    
    End If
 
Next
    
For I = 1 To a

Worksheets(3).Activate
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

Next

Printvalue = 2

    For I = 2 To 705714

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        Label = Cells(I, 1).Value

        TotalVol = TotalVol + Cells(I, 7).Value

        Range("I" & Printvalue).Value = Label
        
        Range("K" & Printvalue).NumberFormat = "0.00"

        Range("L" & Printvalue).Value = TotalVol

        Printvalue = Printvalue + 1

        TotalVol = 0

    Else

        TotalVol = TotalVol + Cells(I, 7).Value

        End If
        
Next
        
OpenValue = Cells(2, 3).Value
Printvalue = 2

    For I = 2 To 705714

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        CloseValue = Cells(I, 6).Value

        YrChange = CloseValue - OpenValue

        Range("J" & Printvalue).Value = YrChange
        
            If YrChange > 0 Then
            
            Range("J" & Printvalue).Interior.ColorIndex = 4
            
            ElseIf YrChange < 0 Then

            Range("J" & Printvalue).Interior.ColorIndex = 3
            
            End If

        PChange = 100 * (CloseValue - OpenValue) / OpenValue
        
        Range("K" & Printvalue).Value = PChange

        Printvalue = Printvalue + 1

        YrChange = 0
    
    End If
    
Next
    
End Sub