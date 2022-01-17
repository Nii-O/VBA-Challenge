Attribute VB_Name = "Module1"
Sub LoopWorksheet()


  Dim Ws As Worksheet
  Application.ScreenUpdating = False
  For Each Ws In Worksheets
    Ws.Select
    Call NextCells
    
  Next
  Application.ScreenUpdating = True
  
End Sub


Sub NextCells()

  ' Set a variable for specifying the column of interest
  Dim column As Integer
  Dim open_count As Double
  Dim close_count As Double
  Dim percent_count As Double
  Dim total_stock As Double
  Dim summary_table_row As Integer
  Dim ticket_name As String
  Dim year_change As Double
  Dim per_change As Double
  Dim WS_count As Integer
  Dim Min As Double
  Dim Max As Double
  Dim MaxV As Double
  Dim Minval As String
  Dim Maxval As String
  Dim MaxVval As String
  
  

  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  column = 1
  open_count = 0
  close_count = 0
  total_stock = 1
  summary_table_row = 2
  WS_count = ActiveWorkbook.Worksheets.Count
  
  
  
      Cells(1, 9).Value = "Ticket"
      Cells(1, 10).Value = "Yearly Change"
      Cells(1, 11).Value = "Percent Change"
      Cells(1, 12).Value = "Total Stock Volume"
      
      
      ' Loop through rows in the column
      For I = 2 To lastrow
    
        ' Searches for when the value of the next cell is different than that of the current cell
        If Cells(I + 1, column).Value <> Cells(I, column).Value Then
    
          ticket_name = Cells(I, 1).Value
          
          open_count = open_count + Cells(I, 3).Value
          close_count = close_count + Cells(I, 6).Value
          total_stock = total_stock + Cells(I, 7).Value
          year_change = close_count - open_count
           If open_count = 0 Then
           per_change = 0
          Else
           per_change = ((close_count - open_count) / open_count) * 100
          End If
          
          Range("I" & summary_table_row).Value = ticket_name
          Range("J" & summary_table_row).Value = year_change
          Range("K" & summary_table_row).Value = per_change
          Range("L" & summary_table_row).Value = total_stock
          
          summary_table_row = summary_table_row + 1
          
          open_count = 0
          close_count = 0
          total_stock = 0
          year_change = 0
          per_change = 0
          
        Else
        
          open_count = open_count + Cells(I, 3).Value
          close_count = close_count + Cells(I, 6).Value
          total_stock = total_stock + Cells(I, 7).Value
          year_change = close_count - open_count
          
          If open_count = 0 Then
           per_change = 0
          Else
           per_change = ((close_count - open_count) / open_count) * 100
          End If
          
        End If
    
      Next I
   
   For c = 2 To summary_table_row
   
    If Cells(c, 10) > 0 Then
        Cells(c, 10).Interior.ColorIndex = 4 'Green
    Else
        Cells(c, 10).Interior.ColorIndex = 3 'Red
    End If
   Next c
   
   
   
   Max = 0
   Maxval = " "
   For v = 2 To summary_table_row
    If Cells(v, 11) > Max Then
        Max = Cells(v, 11)
        Maxval = Cells(v, 9)
    End If
    Cells(2, 17).Value = Max
    Cells(2, 16).Value = Maxval
   Next v
   
   
  Min = 0
  Minval = " "
  For b = 2 To summary_table_row
    If Cells(b, 11) < Min Then
        Min = Cells(b, 11)
        Minval = Cells(b, 9)
    End If
    Cells(3, 17).Value = Min
    Cells(3, 16).Value = Minval
   Next b
   
   
   MaxV = 0
   MaxVval = " "
   For k = 2 To summary_table_row
    If Cells(k, 12) > MaxV Then
        MaxV = Cells(k, 12)
        MaxVval = Cells(k, 9)
    End If
    Cells(4, 17).Value = MaxV
    Cells(4, 16).Value = MaxVval
   Next k
   
    
   Cells(2, 15).Value = "Greatest % Increase"
   Cells(3, 15).Value = "Greatest % Decrease"
   Cells(4, 15).Value = "Greatest Total Volume"
   Cells(1, 16).Value = "Ticker"
   Cells(1, 17).Value = "Value"
   Range("K:K").NumberFormat = "0.00%"
   Range("Q2:Q3").NumberFormat = "0.00%"
   
End Sub


