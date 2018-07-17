VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call VMFileGetter
End Sub

Private Sub CommandButton2_Click()
    Worksheets("Paste Data Here").Cells.Clear 'Clear the worksheet for the next match

End Sub




