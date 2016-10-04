Attribute VB_Name = "expimp"
Option Explicit
Enum FMT
  SEC = 0
  siap_compra = 1
  SIAP_VENTA = 2
End Enum

Public Sub exportar(rec As ADODB.Recordset, archivo As String, formato As FMT)
  Dim i As Integer, j As Integer
  Dim line As String
  Open "csv/" & archivo & IIf(formato = SEC, ".csv", ".txt") For Output As #1
  Select Case formato
  Case SEC
    Do Until rec.EOF
      line = ""
      For j = 0 To rec.fields.Count - 1
        line = line & "@" & IIf(IsNull(rec.fields(j)), "", rec.fields(j))
      Next
      Print #1, Mid(line, 2)
      rec.MoveNext
    Loop
  Case Else
    Dim cod As String, cuit As String, fecha As String
    Do Until rec.EOF
      Print #1, Format(rec!fecha, "yyyymmdd");
      Print #1, Format(query("comprobantes", , "cod_comp=" & rec!letra)!cod_comp_siap, String(3, "0"));
      Print #1, Format(rec!sucursal, String(5, "0"));
      Print #1, Format(rec!n_comp, String(20, "0"));
      If formato = SIAP_VENTA Then Print #1, Format(rec!n_comp, String(20, "0")); Else Print #1, String(16, " ");
      With query("clientes", , "cod_cli=" & rec!cod_cli)
        Print #1, IIf(IsNull(!cuit_cli), "99", "80");
      End With
      If formato = siap_compra Then
        With query("proveedores", , "cod_prov=" & rec!cod_prov)
          Print #1, Format(!cuit_prov, String(20, "0"));
          Print #1, left2(!nom_prov, 30);
        End With
      Else
        With query("clientes", , "cod_cli=" & rec!cod_cli)
          Print #1, left2(coalesce(!cuit_cli, "CONSUMIDOR FINAL"), 20);
          Print #1, left2(!nom_cli, 30);
        End With
      End If
      If formato = siap_compra Then
        Print #1, Format(100 * (rec!no_gravado + rec!gravado + rec!iva21 + rec!iva105 + rec!iva27 + _
                                rec!exento + rec!perc_iva + rec!perc_ib), String(15, "0"));
      Else
        Print #1, Format(100 * (rec!no_gravado + rec!gravado + rec!iva21 + rec!iva105 + rec!iva27 + _
                                rec!exento + rec!ret_iva + rec!ret_ib), String(15, "0"));
      End If
      Print #1, Format(100 * rec!no_gravado, String(15, "0"));
      If formato = SIAP_VENTA Then Print #1, String(15, "0");
      Print #1, Format(100 * rec!exento, String(15, "0"));
      If formato = siap_compra Then
        Print #1, Format(100 * rec!perc_iva, String(15, "0"));
      Else
        Print #1, Format(100 * rec!ret_iva, String(15, "0"));
      End If
      If formato = siap_compra Then Print #1, String(15, "0");
      If formato = siap_compra Then
        Print #1, Format(100 * rec!perc_ib, String(15, "0"));
      Else
        Print #1, Format(100 * rec!ret_ib, String(15, "0"));
      End If
      Print #1, String(15, "0");
      Print #1, Format(100 * rec!interno, String(15, "0"));
      Print #1, "PES";
      Print #1, "0001000000";
      Print #1, Format(IIf(rec!iva21 > 0, 1, 0) + IIf(rec!iva105 > 0, 1, 0) + IIf(rec!iva27 > 0, 1, 0), "0");
      Print #1, "0";
      Print #1, String(15, "0");
      If formato = siap_compra Then
        Print #1, String(15, "0");
        Print #1, String(11, "0");
        Print #1, String(30, " ");
        Print #1, String(15, "0")
      Else
        Print #1, Format(rec!fecha, "yyyymmdd")
      End If
      rec.MoveNext
    Loop
  End Select
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
      For i = 0 To UBound(camp): .fields(i) = IIf(camp(i) = "", Null, camp(i)): Next
      .Update
    Loop
  End With
  Close #1
End Sub
