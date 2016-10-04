VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicio"
   ClientHeight    =   7590
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12435
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar gstatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   7215
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   3360
      ScaleHeight     =   6945
      ScaleWidth      =   8985
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
   Begin VB.CommandButton cmdvolver 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin MSComctlLib.TreeView trinicio 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   11456
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   90
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdvolver_Click()
  Dim f As Form
  If Not hijo Is Nothing Then Unload hijo
  Set StatusBar1 = Nothing
  For Each f In Forms: Unload f: Next
  login.Show
End Sub

Private Sub Form_Load()
  Dim pusu As Boolean, pact As Boolean, pecp As Boolean, pcom As Boolean
  Dim plis As Boolean, ppla As Boolean, pcue As Boolean
  centrar Me
  Set StatusBar1 = gstatus
  pusu = p And 2 ^ 0
  pact = p And 2 ^ 1
  pecp = p And 2 ^ 2
  pcom = p And 2 ^ 3
  plis = p And 2 ^ 4
  ppla = p And 2 ^ 5
  pcue = p And 2 ^ 6

  With trinicio.Nodes
    trinicio.Nodes.Add , , "t0", "ALTAS"
      If pact Then .Add "t0", tvwChild, "t00", "ACTIVIDADES"
      If pecp Then .Add "t0", tvwChild, "t01", "CLIENTES"
      If pcom Then .Add "t0", tvwChild, "t02", "COMPRAS / GASTOS"
      If pcue Then .Add "t0", tvwChild, "t03", "CUENTAS"
      If pecp Then .Add "t0", tvwChild, "t04", "EMPRESAS"
      If pecp Then .Add "t0", tvwChild, "t06", "PROVEEDORES"
      If pusu Then .Add "t0", tvwChild, "t07", "USUARIOS"
      If pcom Then .Add "t0", tvwChild, "t08", "VENTAS / COBROS"
    .Add , , "t1", "BAJAS / MODIFICACIONES"
      If pecp Then .Add "t1", tvwChild, "t10", "CLIENTES"
      If pcom Then .Add "t1", tvwChild, "t11", "COMPRAS / GASTOS"
      If pcue Then .Add "t1", tvwChild, "t12", "CUENTAS"
      If pecp Then .Add "t1", tvwChild, "t13", "EMPRESAS"
      If pecp Then .Add "t1", tvwChild, "t14", "PROVEEDORES"
      If pusu Then .Add "t1", tvwChild, "t15", "USUARIOS"
      If pcom Then .Add "t1", tvwChild, "t16", "VENTAS / COBROS"
    .Add , , "t7", "CONSULTAS"
      .Add "t7", tvwChild, "t70", "CLIENTES"
      .Add "t7", tvwChild, "t71", "EMPRESAS"
      .Add "t7", tvwChild, "t72", "PROVEEDORES"
    .Add , , "t2", "LISTADOS"
    If plis Then
      .Add "t2", tvwChild, "t20", "COMBUSTIBLE"
      .Add "t2", tvwChild, "t21", "IVA COMPRAS"
      .Add "t2", tvwChild, "t22", "IVA VENTAS"
      .Add "t2", tvwChild, "t23", "PERC.ING.BRUTOS"
    End If
    If ppla Then
      .Add , , "t5", "PLANILLA LIQUIDACIÓN"
      .Add "t5", tvwChild, "t50", "INGRESOS BRUTOS"
      .Add "t5", tvwChild, "t51", "IVA"
    End If
    .Add , , "t4", "CAMBIAR CONTRASEÑA"
    If pcom Then
      .Add , , "t6", "TOTALES / PERIODOS"
      .Add , , "t3", "CONTROL DE DUPLICIDAD"
      .Add , , "t8", "IMPORTAR / EXPORTAR"
    End If
  End With
End Sub

Private Sub trinicio_NodeClick(ByVal Node As Node)
  Select Case Node.key
  Case "t00": abrir Frame1, aactividad
  Case "t01": abmcliente.alta = True: abrir Frame1, abmcliente
  Case "t02": abmegreso.alta = True: abrir Frame1, abmegreso
  Case "t03": abrir Frame1, abmcuenta
  Case "t04": abmempresa.alta = True: abrir Frame1, abmempresa
  Case "t06": abmproveedor.alta = True: abrir Frame1, abmproveedor
  Case "t07": abmusuario.alta = True: abrir Frame1, abmusuario
  Case "t08": abmingreso.alta = True: abrir Frame1, abmingreso
  Case "t10": abmcliente.alta = False: abrir Frame1, abmcliente
  Case "t11": abmegreso.alta = False: abrir Frame1, abmegreso
  Case "t12": abrir Frame1, abmcuenta
  Case "t13": abmempresa.alta = False: abrir Frame1, abmempresa
  Case "t14": abmproveedor.alta = False: abrir Frame1, abmproveedor
  Case "t15": abmusuario.alta = False: abrir Frame1, abmusuario
  Case "t16": abmingreso.alta = False: abrir Frame1, abmingreso
  Case "t20": abrir Frame1, gcombustible
  Case "t21": abrir Frame1, givacompras
  Case "t22": abrir Frame1, givaventas
  Case "t23": abrir Frame1, lingbrutos
  Case "t3": abrir Frame1, duplicidad
  Case "t4": abrir Frame1, mclave
  Case "t50": plingbrutos.Show vbModal
  Case "t51": pliva.Show vbModal
  Case "t6": abrir Frame1, calculo
  Case "t8": abrir Frame1, impexp
  Case "t70": buscard.excla = ""
    formbuscard Frame1, "clientes", _
      "nom_cli", "cod_cli", "nom_cli", _
      "cod_cli as Código|format(cuit_cli,'00-00000000-0') as CUIT|nom_cli as Razón_social"
  Case "t72": buscard.excla = ""
    formbuscard Frame1, "proveedores", _
      "nom_prov", "cod_prov", "nom_prov", _
      "cod_prov as Código|format(cuit_prov,'00-00000000-0') as CUIT|nom_prov as Razón_social"
  Case "t71":
    buscard.excla = "cod_emp"
    buscard.excol = "nom_act"
    buscard.extab = "emp_act as ea inner join actividades as a on ea.cod_act=a.cod_act"
    formbuscard Frame1, "empresas", _
      "nom_emp", "cod_emp", "nom_emp", _
      "cod_emp as Código|format(cuit_emp,'00-00000000-0') as CUIT|nom_emp as Razón_social|" & _
      "dom_emp as Domicilio|loc_emp as Localidad|tel_emp as Teléfono|sus_emp as Suscribe|" & _
      "car_emp as Carácter|resp_emp as Responsabilidad_IVA"
  End Select
End Sub
