Attribute VB_Name = "comteclado"
Public Function teclaemp(ByRef emp As TextBox, ByRef lab As Object) As Boolean
  formbuscar "empresas", "nom_emp", "cod_emp", "nom_emp"
  If Not buscar.Cancel Then
    emp = buscar.key
    lab = buscar.val
    crearegresos CStr(emp): crearingresos CStr(emp)
  End If
  teclaemp = Not buscar.Cancel
End Function

Public Function teclacue(ByRef cue As TextBox, ByVal emp As String, ByRef lab As Object) As Boolean
  assert emp <> "", NOCAMP, "Falta ingresar empresa"
  formbuscar2 emp, "cuentas", "nom_cue", "cod_cue", "cod_pad"
  If Not buscar.Cancel Then
    cue = buscar2.key
    lab = buscar2.val
  End If
  teclacue = Not buscar.Cancel
End Function

Public Function teclaprov(ByRef prov As TextBox, ByRef lab As Object) As Boolean
  formbuscar "proveedores", "nom_prov", "cod_prov", "nom_prov"
  If Not buscar.Cancel Then
    prov = buscar.key
    lab = buscar.val
  End If
  teclaprov = Not buscar.Cancel
End Function

Public Function teclacli(ByRef cli As TextBox, ByRef lab As Object) As Boolean
  formbuscar "clientes", "nom_cli", "cod_cli", "nom_cli"
  If Not buscar.Cancel Then
    cli = buscar.key
    lab = buscar.val
  End If
  teclacli = Not buscar.Cancel
End Function

Public Sub teclacue1()
  acuenta.tmp = True
  abrir inicio.Frame1, acuenta, False
End Sub

Public Sub teclaprov1()
  abmproveedor.tmp = True
  abmproveedor.alta = True
  abrir inicio.Frame1, abmproveedor, False
End Sub

Public Sub teclacli1()
  abmcliente.tmp = True
  abmcliente.alta = True
  abrir inicio.Frame1, abmcliente, False
End Sub
