VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form gcombustible 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtemp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8520
      Picture         =   "gcombustible.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   14
      Top             =   6000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4080
      Picture         =   "gcombustible.frx":061A
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   13
      Top             =   6000
      Width           =   375
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
      Left            =   6480
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   8040
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   8520
      Top             =   120
   End
   Begin VB.TextBox txtbuscarcue 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   6000
      Width           =   4095
   End
   Begin VB.TextBox txtbuscarprov 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   3975
   End
   Begin MSComctlLib.TreeView trcuentas 
      Height          =   5175
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9128
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
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
   Begin MSComctlLib.ListView lstproveedores 
      Height          =   5175
      Left            =   120
      TabIndex        =   6
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
      Left            =   1800
      TabIndex        =   3
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
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
      Left            =   3960
      TabIndex        =   4
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
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
   Begin MSComctlLib.ListView lstproveedores1 
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
         Name            =   "Terminus"
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
      Left            =   7320
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labnom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3525
      TabIndex        =   16
      Top             =   240
      Width           =   4815
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
      Left            =   645
      TabIndex        =   15
      Top             =   360
      Width           =   975
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
      Left            =   960
      TabIndex        =   12
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha:"
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
      TabIndex        =   9
      Top             =   6600
      Width           =   855
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
      Left            =   3240
      TabIndex        =   8
      Top             =   6600
      Width           =   735
   End
End
Attribute VB_Name = "gcombustible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgenerar_Click()
  On Error GoTo E
  assert txtemp <> "", NOCAMP, "Falta ingresar empresa"
  crearingresos txtemp: crearegresos txtemp
  Dim cprov As New Collection
  'cadena sql para proveedores
  For Each i In lstproveedores.ListItems
    If i.Checked Then cprov.Add Mid(i.Tag, 2)
  Next
  sprov = borden(cjoin(cprov, ","), "(", ")")
  'cadena sql para cuentas
  Dim ccue As New Collection
  For Each j In trcuentas.Nodes
    h = busc("select n_hijos from cuentas where cod_cue=" & Mid(j.key, 2))!n_hijos
    If h = 0 And j.Checked Then ccue.Add Mid(j.key, 2)
  Next
  scue = borden(cjoin(ccue, ","), "(", ")")
  'cadena sql para creacion de consulta
  If viewexiste("vegresos") Then C.Execute "drop view vegresos"
  selecte = "select cod_egr,fecha,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero,nom_prov,format(cuit_prov,'00-00000000-0') as cuit_prov,nom_cue,litros,format(litros*0.27,'0.00') as litros27 " & _
            "from ((egresos" & txtemp & " as e " & _
            "inner join cuentas as cu on cu.cod_cue=e.cod_cue) " & _
            "inner join proveedores as pr on pr.cod_prov=e.cod_prov) " & _
            "where litros>0"
  If sprov <> "" Then selecte = selecte & " and e.cod_prov in " & sprov
  If scue <> "" Then selecte = selecte & " and e.cod_cue in " & scue
  If txtfecha(0) <> "  /  /    " Then selecte = selecte & " and fecha>=#" & Format(txtfecha(0), "mm/dd/yyyy") & "#"
  If txtfecha(1) <> "  /  /    " Then selecte = selecte & " and fecha<=#" & Format(txtfecha(1), "mm/dd/yyyy") & "#"
  C.Execute "create view vegresos  as " & selecte
  lcombustible.Show vbModal
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  'se tienen 2 listas para mantener mantener los items tildados
  initlst lstproveedores, Array("Proveedores"), Array(0.95)
  initlst lstproveedores1, Array("Proveedores"), Array(0.95)
  llenarlst lstproveedores, "select * from proveedores", Array("nom_prov"), "cod_prov"
  llenarlst lstproveedores1, "select * from proveedores", Array("nom_prov"), "cod_prov"
End Sub

Private Sub lstproveedores_Click()
  lstproveedores1.ListItems(lstproveedores.SelectedItem.Index).Checked = lstproveedores.SelectedItem.Checked
End Sub

Private Sub Timer3_Timer()
  Timer3.Enabled = False
  trcuentas.Nodes.Clear
  For Each n In trcuentas1.Nodes
    If n.Children > 0 Or InStr(1, n, txtbuscarcue, 1) > 0 Then
      If n.Parent Is Nothing Then trcuentas.Nodes.Add , , n.key, n Else trcuentas.Nodes.Add n.Parent.key, tvwChild, n.key, n
    End If
  Next
  For Each i In trcuentas.Nodes: i.Checked = trcuentas1.Nodes(i.key).Checked: Next
End Sub

Private Sub trcuentas_NodeCheck(ByVal Node As Node)
  tildarAbajo Node
  tildarArriba Node
  For Each i In trcuentas.Nodes: trcuentas1.Nodes(i.key).Checked = i.Checked: Next
End Sub

Private Sub txtbuscarprov_Change()
  Timer2.Enabled = True
  Timer2.Interval = 500
End Sub

Private Sub txtbuscarcue_Change()
  Timer3.Enabled = True
  Timer3.Interval = 500
End Sub

Private Sub Timer2_Timer()
  llenarlst lstproveedores, "select * from proveedores where nom_prov like '%" & txtbuscarprov & "%'", Array("nom_prov"), "cod_prov"
  Timer2.Enabled = False
  For Each i In lstproveedores.ListItems: i.Checked = lstproveedores1.ListItems(i.Index).Checked: Next
End Sub

Private Sub txtemp_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de empresa. F3: buscar"
End Sub

Private Sub txtemp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    If teclaemp(txtemp, labnom) Then cargar
  End If
End Sub

Private Sub txtemp_Validate(Cancel As Boolean)
  If txtemp <> "" Then Cancel = validaremp(txtemp, labnom) Else labnom = ""
  If Not Cancel And txtemp <> "" Then cargar
End Sub

Private Sub txtemp_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtfecha_GotFocus(Index As Integer)
  txtfecha(Index).SelStart = 0
  txtfecha(Index).SelLength = 10
End Sub

Private Sub cargar()
  crearingresos txtemp: crearegresos txtemp
  With busc("select nom_emp from empresas where cod_emp=" & txtemp)
    llenarNivel trcuentas, "select * from cuentas where n_hijos>0", "nom_cue", "cod_cue", "cod_pad"
    llenarNivel trcuentas, "select emp_cue.cod_cue,emp_cue.cod_emp,cuentas.nom_cue,cuentas.cod_pad " & _
                             "from emp_cue inner join cuentas on emp_cue.cod_cue=cuentas.cod_cue " & _
                             "where emp_cue.cod_emp=" & txtemp, _
                             "nom_cue", "cod_cue", "cod_pad", False
    'se tienen 2 arboles para la busqueda
    llenarNivel trcuentas1, "select * from cuentas where n_hijos>0", "nom_cue", "cod_cue", "cod_pad"
    llenarNivel trcuentas1, "select emp_cue.cod_cue,emp_cue.cod_emp,cuentas.nom_cue,cuentas.cod_pad " & _
                              "from emp_cue inner join cuentas on emp_cue.cod_cue=cuentas.cod_cue " & _
                              "where emp_cue.cod_emp=" & txtemp, _
                              "nom_cue", "cod_cue", "cod_pad", False
  End With
End Sub

Private Sub txtfecha_Validate(Index As Integer, Cancel As Boolean)
  If txtfecha(Index) <> "  /  /    " Then Cancel = Not validarfecha(txtfecha(Index))
End Sub
