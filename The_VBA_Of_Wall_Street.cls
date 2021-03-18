Attribute VB_Name = "Module1"
Sub WallStreetVBA()

'Name of my variables

Dim mname, miname, gtname As String
Dim n, i, t, nw, w As Integer
Dim c, gt As LongLong
Dim o, cl, p, max, min As Double
Dim counter As Long

'nw = number of worksheets in my workbook
nw = Worksheets.Count

'For loop to loop from w(number of worksheet) to nw(number of worksheets)
For w = 1 To nw

'Activating the actual worksheet according to the loop
Worksheets(w).Activate
t = 0
c = 0

'n = number of rows
n = WorksheetFunction.CountA(Columns("A:A"))

'For loop to loop across all the rows, i(actual row) to n(total number of rows)
For i = 1 To n
counter = counter + 1
    If Range("A" & 1 + i) <> Range("A" & 2 + i) Then
        t = t + 1
            If t > 1 Then
                counter = i - counter
                o = Range("C" & counter + 2)
                cl = Range("F" & i + 1)
                counter = 0
            Else
                o = Range("C2")
                cl = Range("F" & 1 + i)
                counter = 0
            End If
            
        c = Range("G" & 1 + i) + c
        Range("J" & 1 + t) = cl - o
        
        'o(open), cl(closing)
        If o = 0 And cl = 0 Then
            Range("K" & 1 + t) = Format(0, "percent")
        ElseIf o = 0 Then
            p = cl
            Range("K" & 1 + t) = Format(p, "percent")
        Else
            p = (cl / o) - 1
            Range("K" & 1 + t) = Format(p, "percent")
        End If
        
        'If biggest saved stock volume(gt) is fewer than the actual cell stock volume, then biggest saved stock volume(gt) = actual cell stock volume
        If gt < Range("L" & 1 + t) Then
            gt = Range("L" & 1 + t)
            gtname = Range("I" & 1 + t)
        Else
            gt = gt
        End If
            
        If min > Range("K" & 1 + t) Then
            min = Range("K" & 1 + t)
            miname = Range("I" & 1 + t)
        Else
            min = min
        End If
        
        If max < Range("K" & 1 + t) Then
            max = Range("K" & 1 + t)
            mname = Range("I" & 1 + t)
        Else
            max = max
        End If
            
        'Establishing color for winners and losers
        If Range("J" & 1 + t) < 0 Then
            Range("J" & 1 + t).Interior.ColorIndex = 3
        ElseIf Range("J" & 1 + t) > 0 Then
            Range("J" & 1 + t).Interior.ColorIndex = 4
        End If
        
        Range("I" & 1 + t) = Range("A" & 1 + i)
        Range("L" & 1 + t) = c
        
        c = 0
    Else
        c = Range("G" & 1 + i) + c
    End If
Next i
    
    'Title/Topics of items
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percentage Change"
    Range("L1") = "Total Stock Volume"

    Range("O1") = "Ticker"
    Range("P1") = "Value"
    Range("N2") = "Greatest % Increase:"
    Range("N3") = "Greatest % Decrease:"
    Range("N4") = "Greatest Total Volume:"
    
    'Format for the min, max, and greatest stock volume
    Range("O2") = mname
    Range("P2") = Format(max, "percent")
    Range("O3") = miname
    Range("P3") = Format(min, "percent")
    Range("O4") = gtname
    Range("P4") = FormatNumber(gt, 0)
    
    'Gives a custom fit to the columns once the data is stored
    Columns("I:I").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").EntireColumn.AutoFit
    Columns("N:N").EntireColumn.AutoFit
    
    'At the end of every year, it is necessary to reset my cumulative values
    max = 0
    min = 0
    gt = 0
    
Next w
    
End Sub


