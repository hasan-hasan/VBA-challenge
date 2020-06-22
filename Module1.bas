Attribute VB_Name = "Module1"
Sub alpha_test2016()


Dim ticker As String
Dim ticker_total As Double
Dim newtable As Integer

 'lastrow = Cells(Rows.Count, 1).End(x1up).Row

ticker_total = 0
newtable = 2

'For i = 2 To lastrow

For i = 2 To 797711

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        ticker = Cells(i, 1).Value
        ticker_total = ticker_total + Cells(i, 7).Value
        Cells(newtable, 12).Value = ticker_total
        Cells(newtable, 9).Value = ticker
        newtable = newtable + 1
        Total = 0
        
    Else
    
    ticker_total = Cells(i, 7).Value
    
    End If
    
    
Next i

End Sub

Sub alpha_test2015()


Dim ticker As String
Dim ticker_total As Double
Dim newtable As Integer

 'lastrow = Cells(Rows.Count, 1).End(x1up).Row

ticker_total = 0
newtable = 2

'For i = 2 To lastrow

For i = 2 To 760192

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        ticker = Cells(i, 1).Value
        ticker_total = ticker_total + Cells(i, 7).Value
        Cells(newtable, 12).Value = ticker_total
        Cells(newtable, 9).Value = ticker
        newtable = newtable + 1
        Total = 0
        
    Else
    
    ticker_total = Cells(i, 7).Value
    
    End If
    
    
Next i

End Sub


Sub alpha_test2014()


Dim ticker As String
Dim ticker_total As Double
Dim newtable As Integer

 'lastrow = Cells(Rows.Count, 1).End(x1up).Row

ticker_total = 0
newtable = 2

'For i = 2 To lastrow

For i = 2 To 705714

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        ticker = Cells(i, 1).Value
        ticker_total = ticker_total + Cells(i, 7).Value
        Cells(newtable, 12).Value = ticker_total
        Cells(newtable, 9).Value = ticker
        newtable = newtable + 1
        Total = 0
        
    Else
    
    ticker_total = Cells(i, 7).Value
    
    End If
    
    
Next i

End Sub

