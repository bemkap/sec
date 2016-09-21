VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form givaventas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4080
      Picture         =   "givaventas.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   17
      Top             =   6000
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8520
      Picture         =   "givaventas.frx":06BA
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   16
      Top             =   6000
      Width           =   375
   End
   Begin Project1.UserControl2 txtbuscarcue 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   6000
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      enabled         =   0   'False
   End
   Begin Project1.UserControl2 txtbuscarcli 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      enabled         =   0   'False
   End
   Begin VB.ComboBox cmbletra 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdgenerar 
      Caption         =   "Generar"
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
      Left            =   7680
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.TreeView trcuentas 
      Height          =   5175
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9128
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Checkboxes      =   -1  'True
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
   Begin MSComctlLib.ListView lstclientes 
      Height          =   5175
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9128
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   0
   End
   Begin MSMask.MaskEdBox txtfecha 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtfecha 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSComctlLib.ListView lstclientes1 
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView trcuentas1 
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.UserControl1 txtemp 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      _extentx        =   2990
      _extenty        =   661
      info            =   "Ingresar código de empresa. F3: buscar"
      tabla           =   "empresas"
      campo           =   "nom_emp"
      clave           =   "cod_emp"
      busq            =   "nom_emp"
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Cód.emp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.Label labnom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   6600
      Width           =   735
   End
End
Attribute VB_Name = "givaventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdgenerar_Click()
  Dim i As ListItem, j As Node, h As Integer
  Dim scli As String, scue As String, selecti As String
  On Error GoTo E
  assert txtemp <> "", NOCAMP, "Falta ingresar emepresa"
  crearingresos txtemp: crearegresos txtemp
  Dim ccli As New Collection
  'cadena sql para clientes
  For Each i In lstclientes.ListItems
    If i.Checked Then ccli.Add Mid(i.tag, 2)
  Next
  scli = borden(cjoin(ccli, ","), "(", ")")
  'cadena sql para cuentas
  Dim ccue As New Collection
  For Each j In trcuentas.Nodes
    h = busc("select n_hijos from cuentas where cod_cue=" & Mid(j.key, 2))!n_hijos
    If h = 0 And j.Checked Then ccue.Add Mid(j.key, 2)
  Next
  scue = borden(cjoin(ccue, ","), "(", ")")
  'cadena sql para creacion de consulta
  If viewexiste("vingresos") Then C.Execute "drop view vingresos"
  selecti = "select cod_ing,fecha,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero," & _
            "nom_comp,nom_cli,format(cuit_cli,'00-00000000-0') as cuit_cli1," & _
            "iif(i.cod_cue is null,'indefinido',nom_cue) as cuenta," & _
            "gravado,no_gravado,iva21,iva105,iva27,exento,ret_iva,ret_ib,interno," & _
            "gravado+no_gravado+format(iva21+iva105+iva27,'0.00')+exento+ret_iva+ret_ib as subtotal " & _
            "from (((ingresos" & txtemp & " as i " & _
            "left join cuentas as cu on cu.cod_cue=i.cod_cue) " & _
            "inner join clientes as cl on cl.cod_cli=i.cod_cli) " & _
            "inner join comprobantes as co on co.cod_comp=i.letra) " & _
            "where true"
  If scli <> "" Then selecti = selecti & " and i.cod_cli in " & scli
  If scue <> "" Then selecti = selecti & " and i.cod_cue in " & scue
  If cmbletra.ListIndex > -1 Then selecti = selecti & " and i.letra=" & cmbletra.ListIndex
  If txtfecha(0) <> "  /  /    " Then selecti = selecti & " and fecha>=#" & Format(txtfecha(0), "mm/dd/yyyy") & "#"
  If txtfecha(1) <> "  /  /    " Then selecti = selecti & " and fecha<=#" & Format(txtfecha(1), "mm/dd/yyyy") & "#"
  C.Execute "create view vingresos as " & selecti
  livaventas.Show vbModal
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  'se tienen 2 listas para mantener mantener los items tildados
  initlst lstclientes, Array("Clientes"), Array(1)
  initlst lstclientes1, Array("Clientes"), Array(1)
  llenarlst lstclientes, "select * from clientes", Array("nom_cli"), "cod_cli"
  llenarlst lstclientes1, "select * from clientes", Array("nom_cli"), "cod_cli"
  llenarcmb cmbletra, "select * from comprobantes", "nom_comp"
End Sub

Private Sub lstclientes_ItemCheck(ByVal item As MSComctlLib.ListItem)
  lstclientes1.ListItems(item.key).Checked = item.Checked
End Sub

Private Sub trcuentas_NodeCheck(ByVal Node As Node)
  Dim i As Node
  tildarabajo Node
  tildararriba Node
  For Each i In trcuentas.Nodes: trcuentas1.Nodes(i.key).Checked = i.Checked: Next
End Sub

Private Sub txtbuscarcli_buscar()
  Dim i As ListItem
  llenarlst lstclientes, "select * from clientes where nom_cli like '%" & txtbuscarcli & "%'", Array("nom_cli"), "cod_cli"
  For Each i In lstclientes.ListItems: i.Checked = lstclientes1.ListItems(i.key).Checked: Next
End Sub

Private Sub txtbuscarcue_buscar()
  Dim n As Node, i As Node
  trcuentas.Nodes.Clear
  For Each n In trcuentas1.Nodes
    If n.Children > 0 Or InStr(1, n, txtbuscarcue, 1) > 0 Then
      If n.Parent Is Nothing Then trcuentas.Nodes.Add , , n.key, n Else trcuentas.Nodes.Add n.Parent.key, tvwChild, n.key, n
    End If
  Next
  For Each i In trcuentas.Nodes: i.Checked = trcuentas1.Nodes(i.key).Checked: Next
End Sub

Private Sub txtemp_finbusqueda(llave As String, valor As String)
  txtemp = llave
  labnom = valor
  cargar
  txtbuscarcli.enabled = True
  txtbuscarcue.enabled = True
  txtfecha(0).enabled = True
  txtfecha(1).enabled = True
  cmbletra.enabled = True
  cmdgenerar.enabled = True
End Sub

Private Sub txtemp_vacio()
  labnom = ""
  txtbuscarcli.enabled = False
  txtbuscarcue.enabled = False
  txtfecha(0).enabled = False
  txtfecha(1).enabled = False
  cmbletra.enabled = False
  cmdgenerar.enabled = False
End Sub

Private Sub txtfecha_GotFocus(Index As Integer)
  txtfecha(Index).SelStart = 0
  txtfecha(Index).SelLength = 10
End Sub

Private Sub cargar()
  crearingresos txtemp: crearegresos txtemp
  With busc("select nom_emp from empresas where cod_emp=" & txtemp)
    llenarnivel trcuentas, "select * from cuentas where n_hijos>0", "nom_cue", "cod_cue", "cod_pad"
    llenarnivel trcuentas, "select emp_cue.cod_cue,emp_cue.cod_emp,cuentas.nom_cue,cuentas.cod_pad " & _
                             "from emp_cue inner join cuentas on emp_cue.cod_cue=cuentas.cod_cue " & _
                             "where emp_cue.cod_emp=" & txtemp, _
                             "nom_cue", "cod_cue", "cod_pad", False
    'se tienen 2 arboles para la busqueda
    llenarnivel trcuentas1, "select * from cuentas where n_hijos>0", "nom_cue", "cod_cue", "cod_pad"
    llenarnivel trcuentas1, "select emp_cue.cod_cue,emp_cue.cod_emp,cuentas.nom_cue,cuentas.cod_pad " & _
                              "from emp_cue inner join cuentas on emp_cue.cod_cue=cuentas.cod_cue " & _
                              "where emp_cue.cod_emp=" & txtemp, _
                              "nom_cue", "cod_cue", "cod_pad", False
  End With
End Sub

Private Sub txtfecha_Validate(Index As Integer, Cancel As Boolean)
  If txtfecha(Index) <> "  /  /    " Then Cancel = Not validarfecha(txtfecha(Index))
End Sub
