Attribute VB_Name = "VBA_Challenge"
Sub VBA_Challenge()

    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
            Dim ticker As String
            Dim TSV As Double
            TSV = 0
            Dim sumtab As Integer
            sumtab = 2
            Dim yearcE As Double
            Dim yearcB As Double
            Dim greatI As Double
            greatI = 0
            Dim greatD As Double
            greatD = 0
            Dim greatTSV As Double
            greatTSV = 0
    
            For i = 2 To 1000000
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    ticker = Cells(i, 1).Value
                    TSV = TSV + Cells(i, 7).Value
                    Range("I" & sumtab).Value = ticker
                    Range("L" & sumtab).Value = TSV
                    sumtab = sumtab + 1
                    TSV = 0
                Else
                    TSV = TSV + Cells(i, 7).Value
                End If
                If Cells(i + 1, 2) < Cells(i, 2).Value Then
                    yearcE = Cells(i, 6).Value
                    yearcB = Cells(i - 250, 3).Value
                    yearchange = yearcE - yearcB
                    Range("J" & sumtab - 1).Value = yearchange
                    Range("K" & sumtab - 1).Value = yearchange / yearcB
                    Range("K" & sumtab - 1).NumberFormat = "0.00%"
                End If
            Next i
            For i = 2 To 5000
                For j = 10 To 10
                    If Cells(i, j) >= 0 Then
                        Cells(i, j).Interior.ColorIndex = 4
                    Else
                        Cells(i, j).Interior.ColorIndex = 3
                    End If
                Next j
            Next i
            For i = 2 To 5000
                For j = 11 To 11
                    If Cells(i, j) >= 0 Then
                        Cells(i, j).Interior.ColorIndex = 4
                    Else
                        Cells(i, j).Interior.ColorIndex = 3
                    End If
                Next j
            Next i
            For i = 2 To 50000
                If Cells(i, 11) > greatI Then
                    greatI = Cells(i, 11).Value
                    Range("Q2").Value = greatI
                    Range("Q2").NumberFormat = "0.00%"
                    Range("P2").Value = Cells(i, 9).Value
                End If
                If Cells(i, 11) < greatD Then
                    greatD = Cells(i, 11).Value
                    Range("Q3").Value = greatD
                    Range("Q3").NumberFormat = "0.00%"
                    Range("P3").Value = Cells(i, 9).Value
                End If
                If Cells(i, 12) > greatTSV Then
                    greatTSV = Cells(i, 12).Value
                    Range("Q4").Value = greatTSV
                    Range("P4").Value = Cells(i, 9).Value
                End If
            Next i
    Next ws
End Sub

