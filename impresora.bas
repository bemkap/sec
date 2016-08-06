Attribute VB_Name = "impresora"
Public Sub centro(ByVal str As String)
  Printer.CurrentX = (Printer.ScaleWidth - Len(str)) / 2
  Printer.Print str;
End Sub

Public Sub derecha(ByVal str As String)
  Printer.CurrentX = Printer.ScaleWidth - Len(str)
  Printer.Print str;
End Sub

Public Sub yx(ByVal y As Integer, ByVal x As Integer, ByVal str As String)
  Printer.CurrentX = x
  Printer.CurrentY = y
  Printer.Print str;
End Sub
