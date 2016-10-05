Attribute VB_Name = "tabla"
Option Explicit
Public C As ADODB.Connection
Public U As String
Public p As Byte

Public Sub crearbd()
  Dim spath As String
  spath = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "db1.mdb;"
  With New ADOX.Catalog
    .Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & spath
  End With
End Sub

Public Function tablaexiste(ByVal tabla As String) As Boolean
  If Not C Is Nothing Then:
    tablaexiste = Not C.OpenSchema(adSchemaTables, Array(Empty, Empty, tabla, "table")).EOF
End Function

Public Function viewexiste(ByVal view As String) As Boolean
  If Not C Is Nothing Then:
    viewexiste = Not C.OpenSchema(adSchemaTables, Array(Empty, Empty, view, "view")).EOF
End Function

Public Sub crearusuarios()
  If Not tablaexiste("usuarios") Then
    Dim Hash As New MD5Hash
    Dim bytBlock() As Byte
    bytBlock = "admin"
    C.Execute "create table usuarios (" & _
                "nombre varchar(30) primary key," & _
                "clave varchar(50)," & _
                "permisos integer)"
    C.Execute "insert into usuarios values ('admin','" & Hash.HashBytes(bytBlock) & "',255)"
  End If
End Sub

Public Sub crearactividades()
  If Not tablaexiste("actividades") Then
    C.Execute "create table actividades (" & _
                "cod_act integer primary key," & _
                "nom_act varchar(200)," & _
                "obs_act varchar(200))"
  End If
  If Not tablaexiste("actividades_cds") Then
    C.Execute "select * into actividades_cds from actividades where 1=0"
  End If
End Sub

Public Sub crearempresas()
  If Not tablaexiste("empresas") Then
    C.Execute "create table empresas (" & _
                "cod_emp identity primary key," & _
                "regvalid bit default true," & _
                "cuit_emp double," & _
                "nom_emp varchar(64)," & _
                "dom_emp varchar(64)," & _
                "loc_emp varchar(64)," & _
                "tel_emp varchar(64)," & _
                "sus_emp varchar(64)," & _
                "car_emp varchar(64)," & _
                "resp_emp varchar(64))"
  End If
End Sub

Public Sub crearempcue()
  crearempresas
  crearcuentas
  If Not tablaexiste("emp_cue") Then
    C.Execute "create table emp_cue (" & _
                "cod_emp integer constraint cue_emp references empresas(cod_emp)," & _
                "cod_cue integer constraint cue_cue references cuentas(cod_cue)," & _
                "constraint emp_cue primary key(cod_emp,cod_cue))"
  End If
End Sub

Public Sub crearempact()
  crearempresas
  crearactividades
  If Not tablaexiste("emp_act") Then
    C.Execute "create table emp_act (" & _
                "cod_emp integer constraint act_emp references empresas(cod_emp)," & _
                "cod_act integer constraint act_act references actividades(cod_act)," & _
                "constraint emp_act primary key(cod_emp,cod_act))"
  End If
End Sub

Public Sub crearcuentas()
  If Not tablaexiste("cuentas") Then
    C.Execute "create table cuentas (" & _
                "cod_cue identity primary key," & _
                "cod_pad integer," & _
                "nom_cue varchar(64)," & _
                "n_hijos integer default 0)"
    C.Execute "insert into cuentas (cod_pad,nom_cue) values (0,'activo')"
    C.Execute "insert into cuentas (cod_pad,nom_cue) values (0,'pasivo')"
    C.Execute "insert into cuentas (cod_pad,nom_cue) values (0,'patrimonio neto')"
    C.Execute "insert into cuentas (cod_pad,nom_cue) values (0,'resultados')"
  End If
End Sub

Public Sub crearingresos(ByVal emp As Integer)
  crearempresas
  crearclientes
  crearcuentas
  If Not tablaexiste("ingresos" & emp) Then
    C.Execute "create table ingresos" & emp & " (" & _
                "cod_ing identity primary key," & _
                "cod_emp integer constraint ing_emp" & emp & " references empresas(cod_emp)," & _
                "sucursal integer," & _
                "n_comp integer," & _
                "fecha date," & _
                "letra integer," & _
                "cod_cli integer constraint ing_cli" & emp & " references clientes(cod_cli)," & _
                "no_gravado double default 0," & _
                "gravado double default 0," & _
                "iva21 double," & _
                "iva105 double," & _
                "iva27 double," & _
                "exento double default 0," & _
                "interno double default 0," & _
                "ret_iva double default 0," & _
                "ret_ib double default 0," & _
                "cod_cue integer constraint ing_cue" & emp & " references cuentas(cod_cue)," & _
                "periodo integer)"
    C.Execute "select * into dingresos" & emp & " from ingresos" & emp & " where 1=2"
  End If
End Sub

Public Sub crearegresos(ByVal emp As Integer)
  crearempresas
  crearproveedores
  crearcuentas
  If Not tablaexiste("egresos" & emp) Then
    C.Execute "create table egresos" & emp & " (" & _
                "cod_egr identity primary key," & _
                "cod_emp integer constraint egr_emp" & emp & " references empresas(cod_emp)," & _
                "sucursal integer," & _
                "n_comp integer," & _
                "fecha datetime," & _
                "letra integer," & _
                "cod_prov integer constraint egr_prov" & emp & " references proveedores(cod_prov)," & _
                "no_gravado double default 0," & _
                "gravado double default 0," & _
                "iva21 double," & _
                "iva105 double," & _
                "iva27 double," & _
                "exento double default 0," & _
                "interno double default 0," & _
                "litros double default 0," & _
                "perc_iva double default 0," & _
                "perc_ib double default 0," & _
                "cod_cue integer constraint egr_cue" & emp & " references cuentas(cod_cue)," & _
                "periodo integer)"
    C.Execute "select * into degresos" & emp & " from egresos" & emp & " where 1=2"
  End If
End Sub

Public Sub crearclientes()
  If Not tablaexiste("clientes") Then
    C.Execute "create table clientes (" & _
                "cod_cli identity primary key," & _
                "regvalid bit default true," & _
                "cuit_cli double," & _
                "nom_cli varchar(64))"
  End If
End Sub

Public Sub crearproveedores()
  If Not tablaexiste("proveedores") Then
    C.Execute "create table proveedores (" & _
                "cod_prov identity primary key," & _
                "regvalid bit default true," & _
                "cuit_prov double," & _
                "nom_prov varchar(64))"
  End If
End Sub

Public Sub crearcomprobantes()
  If Not tablaexiste("comprobantes") Then
    C.Execute "create table comprobantes (" & _
                "cod_comp identity primary key," & _
                "cod_comp_siap integer," & _
                "nom_comp varchar(5)," & _
                "ivadisc_comp bit)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (0,'FAC A',1,True)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (1,'FAC B',6,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (2,'FAC C',11,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (3,'REC A',4,True)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (4,'REC B',9,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (5,'REC C',15,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (6,'NCR A',3,True)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (7,'NCR B',8,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (8,'NCR C',13,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (9,'NDB A',2,True)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (10,'NDB B',7,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (11,'NDB C',12,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (12,'TIC Z',83,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (13,'TIC A',81,True)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (14,'TIC B',82,False)"
    C.Execute "insert into comprobantes (cod_comp,nom_comp,cod_comp_siap,ivadisc_comp) values (15,'TIC C',83,False)"
  End If
End Sub
