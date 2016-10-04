Attribute VB_Name = "registro"
Option Explicit

Public Function exec(sql As String) As ADODB.Recordset
  Set exec = New ADODB.Recordset
  exec.CursorLocation = adUseClient
  exec.Open sql, C, adOpenStatic, adLockOptimistic, adCmdText
End Function

Public Function tabl(tbl As String) As ADODB.Recordset
  Set tabl = New ADODB.Recordset
  tabl.CursorLocation = adUseClient
  tabl.Open tbl, C, adOpenStatic, adLockOptimistic, adCmdTable
End Function

Public Function query(ByVal from As String, Optional ByVal fields As String = "*", Optional ByVal where As String = "", Optional ByVal orderby As String = "", Optional ByVal groupby As String = "") As ADODB.Recordset
  Dim s As String
  Set query = New ADODB.Recordset
  query.CursorLocation = adUseClient
  s = "select " & fields & " from " & from
  If where <> "" Then s = s & " where " & where
  If orderby <> "" Then s = s & " order by " & orderby
  If groupby <> "" Then s = s & " group by " & groupby
  query.Open s, C, adOpenStatic, adLockOptimistic, adCmdText
End Function

Public Sub llenarcmb(cmb As ComboBox, sql As String, fnom As String, Optional fdat As String = "")
  cmb.Clear
  With exec(sql)
    Do Until .EOF
      cmb.AddItem .fields(fnom)
      If fdat <> "" Then cmb.ItemData(cmb.NewIndex) = .fields(fdat)
      .MoveNext
    Loop
  End With
End Sub

Public Sub llenarlst(lst As ListView, sql As String, campo As Variant, Optional ByVal llave As String = "", Optional vaciar As Boolean = True, Optional pref As String = "k")
  Dim i As Integer, n As Integer, k As String
  If vaciar Then lst.ListItems.Clear
  With exec(sql)
    Do Until .EOF
      Dim item As ListItem
      If llave = "" Then k = pref & n Else k = pref & .fields(llave)
      n = n + 1
      Set item = lst.ListItems.Add(, k, .fields(ascampo(campo(0))))
      item.tag = k
      For i = 1 To UBound(campo)
        If campo(i) <> "" And Not IsNull(.fields(ascampo(campo(i)))) Then
          item.ListSubItems.Add , , .fields(ascampo(campo(i)))
        Else
          item.ListSubItems.Add , , ""
        End If
      Next
      .MoveNext
    Loop
  End With
End Sub

Public Sub initlst(lst As ListView, col As Variant, anc As Variant)
  Dim i As Integer
  For i = 0 To UBound(col)
    With lst.ColumnHeaders.Add()
      .text = ascampo(col(i))
      .Width = anc(i) * lst.Width
    End With
  Next
End Sub

Public Sub formbuscar(tabla As String, campo As String, clave As String, busq As String, valido As String)
  buscar.tabla = tabla
  buscar.columna = campo
  buscar.clave = clave
  buscar.busq = busq
  buscar.valido = valido
  buscar.Show vbModal
End Sub

Public Sub formbuscar2(codemp As String, tabla As String, campo As String, clave As String, padre As String)
  buscar2.codemp = codemp
  buscar2.tabla = tabla
  buscar2.columna = campo
  buscar2.clave = clave
  buscar2.padre = padre
  buscar2.Show vbModal
End Sub
