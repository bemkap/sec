Attribute VB_Name = "comvalidar"
Option Explicit

Public Function validarfecha(fecha As String) As Boolean
  Dim t As Date
  On Error GoTo E
  t = CDate(fecha)
  StatusBar1.SimpleText = ""
  validarfecha = True
  Exit Function
E:
  StatusBar1.SimpleText = "Fecha incorrecta"
  validarfecha = False
End Function
