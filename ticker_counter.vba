Sub total_volume_per_symbol()

'Declare variable called row as data type double
Dim row As Double
'Declare variable called lastrow as data type double
Dim lastrow As Double
'Declare variable called ticker as data type string
Dim ticker As String
'Declare variable called nextticker as data type string
Dim nextticker As String
'Declare variable called totalrow as data type integer
Dim currenttotalrow As Integer
'Declare variable called volume_counter as data type variant
Dim volume_counter As Variant

'initialize the variable currenttotalrow as 2
currenttotalrow = 2
'initialize the variable volume_counter as 0
volume_counter = 0

'populate header row with appropriate name for column
Cells(1, 9).Value = "<ticker>"
'populate header row with appropriate name for column
Cells(1, 10).Value = "<ttl_volume>"

'initialize variable lastrow as one row up from first row with empty cell
lastrow = Cells(Rows.Count, 1).End(xlUp).row

'begin For loop to scan rows
For row = 2 To lastrow
'initialize variable ticker to track which stock is being scanned
ticker = Cells(row, 1).Value
 'initialize variable nextticker to monitor current ticker relative to next ticker
 nextticker = Cells(row + 1, 1).Value
 'concatenate values for each cell with volume of daily trades
 volume_counter = volume_counter + Cells(row, 7).Value
 
 'If block telling For loop to move on to next ticker
 If (ticker <> nextticker) Then
  Cells(currenttotalrow, 9).Value = ticker
  Cells(currenttotalrow, 10).Value = volume_counter
  volume_counter = 0
  currenttotalrow = currenttotalrow + 1
 End If
Next

End Sub
