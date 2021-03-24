Sub stockvalue()

Dim ws As Worksheet
    
   For Each ws In Worksheets
    ws.Activate
           
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
  
Dim ticker As String

Dim totalvolume As Single
totalvolume = 0

Dim Summary As Integer
Summary = 2

Dim opened As Double
opened = Cells(2, 3).Value

Dim closed As Double

Dim finalvalue As Double

Dim growthpercentage As Double

lastrow = Cells(Rows.Count, 1).End(xlUp).Row


    For x = 2 To lastrow
    
            If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
                    
                ticker = Cells(x, 1).Value
    
                totalvolume = totalvolume + Cells(x, 7).Value
                
                closed = Cells(x, 6).Value
                
                finalvalue = closed - opened

                    If finalvalue < 0 Then
                        Range("j" & Summary).Interior.ColorIndex = 3
                        
                    Else
                        Range("j" & Summary).Interior.ColorIndex = 4
                     End If
                
                    growthpercentage = (closed / opened) - 1
                    
                Range("K" & Summary).Value = growthpercentage
                     Range("k" & Summary).NumberFormat = "0.00%"
                     
                opened = Cells(x + 1, 3).Value
                
                If opened = 0 Then
                
                    opened = Cells(x + 1, 3).End(xlDown)
                    
                  End If
                  
                
                Range("I" & Summary).Value = ticker
            
                Range("L" & Summary).Value = totalvolume
                
                Range("J" & Summary).Value = finalvalue
               
                Summary = Summary + 1
                  
                totalvolume = 0
                    
                Else
            
                totalvolume = totalvolume + Cells(x, 7).Value
             
            End If
             
    Next x
    
 

  Next ws



End Sub