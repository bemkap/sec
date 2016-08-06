Attribute VB_Name = "comvalidar"
Public Function validaremp(ByVal emp As String, ByRef lab As Object) As Boolean
  With busc("select * from empresas where cod_emp=" & emp)
    If .RecordCount > 0 Then
      lab = !nom_emp
      StatusBar1.SimpleText = ""
      crearegresos CStr(emp): crearingresos CStr(emp)
    Else
      StatusBar1.SimpleText = "Empresa inexistente"
      lab = ""
      validaremp = True
    End If
  End With
End Function

Public Function validarcue(ByVal cue As String, ByVal emp As String, ByRef lab As Object) As Boolean
  assert emp <> "", NOCAMP, "Ingresar empresa"
  With busc("select * from cuentas where cod_cue=" & cue)
    If .RecordCount > 0 Then
      If !n_hijos > 0 Then
        StatusBar1.SimpleText = "La cuenta no es usable"
        validarcue = True
      ElseIf busc("select * from emp_cue where cod_cue=" & cue & " and cod_emp=" & emp).RecordCount = 0 Then
        StatusBar1.SimpleText = "La cuenta no está incluída en el plan de cuentas"
        validarcue = True
      Else
        lab = !nom_cue
        StatusBar1.SimpleText = ""
      End If
    Else
      StatusBar1.SimpleText = "Cuenta inexistente"
      lab = ""
      validarcue = True
    End If
  End With
End Function

Public Function validarprov(ByVal prov As String, ByRef lab As Object) As Boolean
  With busc("select * from proveedores where cod_prov=" & prov)
    If .RecordCount > 0 Then
      lab = !nom_prov
      StatusBar1.SimpleText = ""
    Else
      StatusBar1.SimpleText = "Proveedor inexistente"
      lab = ""
      validarprov = True
    End If
  End With
End Function

Public Function validarcli(ByVal cli As String, ByRef lab As Object) As Boolean
  With busc("select * from clientes where cod_cli=" & cli)
    If .RecordCount > 0 Then
      lab = !nom_cli
      StatusBar1.SimpleText = ""
    Else
      StatusBar1.SimpleText = "Cliente inexistente"
      lab = ""
      validarcli = True
    End If
  End With
End Function

Public Function validarfecha(fecha As String) As Boolean
  On Error GoTo E
  t = CDate(fecha)
  StatusBar1.SimpleText = ""
  validarfecha = True
  Exit Function
E:
  StatusBar1.SimpleText = "Fecha incorrecta"
  validarfecha = False
End Function
