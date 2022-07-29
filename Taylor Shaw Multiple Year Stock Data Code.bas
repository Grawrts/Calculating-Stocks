Attribute VB_Name = "Module1"
Sub stocks():

  ' Set an initial variable for holding the ticker symbol
  Dim ticker As String
  Dim result As Variant
  Dim volume As Double
  Dim difference As Double

j = 2
  

 startrow = 2
    ' Loop through all stock
  For i = 2 To Sheet1.Cells(Rows.Count, "A").End(xlUp).Row
  

ticker = Cells(i, 1).Value

     ' Check if we are still within the same ticker symbol, if we are not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Write ticker symbol in ticker column
    Cells(j, 9) = ticker
    
    'Subtract beginning open price from end closing price to get yearly change
    difference = Cells(i, 6).Value - Cells(startrow, 3)
    
    Cells(j, 10) = difference
    
   'Calculate percentage change from beginning open price to end closing price
    result = difference / Cells(startrow, 3)
    
    'Insert result into Cells(j,11).Value
    Cells(j, 11).Value = result
    
    'Sum first row of volume to last row of volume
    volume = volume + Cells(i, 7).Value
    
    Cells(j, 12) = volume
    
    'reset variables
    ticker = 0
    result = 0
    volume = 0
    startrow = i + 1
    difference = 0
    j = j + 1
    
    Else
    
    volume = volume + Cells(i, 7).Value
    
        'If Cells(j, 10).Value < 0 Then
            'Cells(j, 10).Interior.ColorIndex = 3
    
       ' Else
            'Cells(j, 10).Interior.ColorIndex = 4
       ' End If


    
    End If
    
     If Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
    
        Else
            Cells(j, 10).Interior.ColorIndex = 4
        End If

    

Next i

End Sub

