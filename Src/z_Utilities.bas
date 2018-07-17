Attribute VB_Name = "z_Utilities"
Function sheettoarray(sheetin As Worksheet) As Variant
    Dim finalarray
    Dim lastcol
    Dim lastrow
    
    lastrow = sheetin.Cells(sheetin.Rows.Count, "A").End(xlUp).Row
    lastcol = sheetin.Cells(1, sheetin.Columns.Count).End(xlToLeft).Column
    
    finalarray = sheetin.Range("A1:" & Number2Letter(lastcol) & lastrow).Value
    
    sheettoarray = finalarray
End Function

Function Number2Letter(colnum As Variant)

Number2Letter = Split(Cells(1, colnum).Address, "$")(1)

End Function


Function asdf()
Application.ScreenUpdating = True
End Function
