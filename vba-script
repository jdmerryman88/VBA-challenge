
Sub stock()

For Each ws In Worksheets

'Set column and row titles
ws.Range("j1").Value = "Ticker"
ws.Range("k1") = "Yearly Change"
ws.Range("l1") = "Percent Change"
ws.Range("l1") = "Total Stock Volume"
ws.Range("p3") = "Greatest % Increase"
ws.Range("p4") = "Greatest % Decrease"
ws.Range("p5") = "Greatest Total Volume"
ws.Range("Q1") = "Ticker"
ws.Range("R1") = "Value"

'Define variables
Dim price As Double
Dim location As Long
Dim symbol As String
Dim final As Double
Dim counter As Long
Dim priceloc As Long
Dim percent As Double
Dim volume As Double
Dim ticker As String
Dim increase As Double
Dim decrease As Double
Dim gvolume As Double

'defines initial values for variables
symbol = ws.Cells(2, 1)
location = 2
counter = 0
volume = 0
ticker = ws.Cells(2, 1)

'determines the last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row



    For i = 2 To lastrow
        
        'check cells to see if same stock is in the next cell
        If ws.Cells(i, 1).Value <> symbol Then
        ticker = ws.Cells(i - 1, 1)
        
        'writes ticker symbol to cell
        ws.Cells(location, 10) = ticker

        'Gets stocks final price
        final = ws.Cells(i - 1, 6)
        
        'Gets stock opening price
        priceloc = i - counter
        price = ws.Cells(priceloc, 3)
        
        'Checks if price of stock is greater than zero
        If price > 0 Then
        
        'Calculates percent change
        percent = (final - price) / price
        
        'Writes yearly change
        ws.Cells(location, 11) = final - price
        
        'Writes percent change
        ws.Cells(location, 12) = percent
        
        'Formats cells
        ws.Cells(location, 12).NumberFormat = "0.00%"
            If ws.Cells(location, 11) > 0 Then
            ws.Cells(location, 11).Interior.ColorIndex = 4
            Else
            ws.Cells(location, 11).Interior.ColorIndex = 3
            End If
         End If
         
        'Writes stocks total volume        
        ws.Cells(location, 13) = volume
        location = location + 1
        counter = 0
        symbol = ws.Cells(i, 1)
        volume = ws.Cells(i, 7)
        End If
        
        counter = counter + 1
        volume = volume + ws.Cells(i, 7)
    Next i
    
    'Creates row labels for greatest increase, decrease and volume change
    increase = ws.Cells(2, 12)
    decrease = ws.Cells(2, 12)
    gvolume = ws.Cells(2, 13)
    
    'determines greatest increase and writes to cell
    For j = 2 To lastrow
        If increase < ws.Cells(j, 12) Then
         increase = ws.Cells(j, 12)
         ws.Cells(3, 17) = ws.Cells(j, 10)
         End If
    Next j
    ws.Cells(3, 18) = increase
    ws.Cells(3, 18).NumberFormat = "0.00%"
    
    'determines greatest decrease and writes to cell
    For k = 2 To lastrow
        If decrease > ws.Cells(k, 12) Then
         decrease = ws.Cells(k, 12)
         ws.Cells(4, 17) = ws.Cells(k, 10)
         End If
    Next k
    ws.Cells(4, 18) = decrease
    ws.Cells(4, 18).NumberFormat = "0.00%"
    
    'determine greatest volume increase and writes to cell
    For l = 2 To lastrow
        If gvolume < ws.Cells(l, 13) Then
         gvolume = ws.Cells(l, 13)
         ws.Cells(5, 17) = ws.Cells(l, 10)
         End If
    Next l
    ws.Cells(5, 18) = gvolume

Next ws

End Sub

