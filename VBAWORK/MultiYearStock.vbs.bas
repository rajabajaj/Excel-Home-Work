Attribute VB_Name = "Module2"
Sub StockVolEasy()


    Dim ws As Worksheet
    Dim Ticker_name As String
    Dim Total_Volume As Double
     
     For Each ws In Worksheets
     Total_Volume = 0
    
    Dim Display_Table As Integer
     Display_Table = 2

    
       
         For i = 2 To 43398

           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              Ticker_name = ws.Cells(i, 1).Value

               Total_Volume = ws.Cells(i, 7).Value

               Total_Volume = Total_Volume + ws.Cells(i, 7).Value

                   ws.Range("J" & Display_Table).Value = Ticker_name
                   ws.Range("K" & Display_Table).Value = Total_Volume
                   ws.Range("J1") = ws.Range("A1").Value
                   ws.Range("K1").Value = ws.Range("G1").Value

                   Display_Table = Display_Table + 1

                   Total_Volume = 0

           Else
           Total_Volume = Total_Volume + ws.Cells(i, 7).Value

        End If

     Next i
     
Next ws

 End Sub
