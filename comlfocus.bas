Attribute VB_Name = "comlfocus"
Public Sub lfocustxtn(ByVal i As Integer, ByRef txtn As Variant, ByRef lablitros As Label, ByRef iva As Double)
  Select Case i
  Case 0: 'gravado
    txtn(4) = val(txtn(0)) * 0.21
  Case 3: 'interno
    txtn(3).Tag = txtn(3)
    txtn(3) = val(txtn(3).Tag) - val(lablitros)
  Case 5: 'litros
    lablitros = Format(val(txtn(5)) * 0.27, "0.00")
    txtn(3) = val(txtn(3).Tag) - val(lablitros)
  Case 4: 'iva
    StatusBar1.SimpleText = ""
  End Select
  iva = txtn(4)
End Sub
