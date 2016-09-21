Attribute VB_Name = "expimp"
Option Explicit

Public Sub exportar(rec As ADODB.Recordset, archivo As String)
  Dim i As Integer, j As Integer
  Dim line As String
  Open "csv/" & archivo & ".csv" For Output As #1
  Do Until rec.EOF
    line = ""
    For j = 0 To rec.Fields.Count - 1
      line = line & "@" & IIf(IsNull(rec.Fields(j)), "", rec.Fields(j))
    Next
    Print #1, Mid(line, 2)
    rec.MoveNext
  Loop
  Close #1
End Sub

Public Sub importar(archivo As String, Optional tabla As String)
  Dim i As Integer
  Dim line As String, camp() As String
  If tabla = "" Then tabla = archivo
  Open archivo For Input As #1
  assert tablaexiste(tabla), INVOP, "La tabla no existe"
  With tabl(tabla)
    Do Until EOF(1)
      Line Input #1, line
      camp = Split(line, "@")
      .AddNew
      For i = 0 To UBound(camp): .Fields(i) = IIf(camp(i) = "", Null, camp(i)): Next
      .Update
    Loop
  End With
  Close #1
End Sub
