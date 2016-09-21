VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form lingbrutos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
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
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstbrutos 
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
   Begin MSComctlLib.ListView lsttotal 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   661
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
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
      Height          =   375
      Index           =   0
      Left            =   1920
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
      Height          =   375
      Index           =   1
      Left            =   4080
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
   Begin Project1.UserControl1 txtemp 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      info            =   "Ingresar código de empresa. F3: buscar"
      tabla           =   "empresas"
      campo           =   "nom_emp"
      clave           =   "cod_emp"
      busq            =   "nom_emp"
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
      Left            =   3240
      TabIndex        =   10
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
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   975
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
      Left            =   3360
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   6600
      Width           =   855
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
      Left            =   1080
      TabIndex        =   6
      Top             =   6600
      Width           =   735
   End
End
Attribute VB_Name = "lingbrutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ws(), n As Double

Private Sub Form_Load()
  initlst lstbrutos, Array("FECHA", "TIPO", "Nº", "RAZÓN SOCIAL", "CUIT", "PERCEPCIÓN"), Array(0.14, 0.1, 0.2, 0.2, 0.2, 0.15)
  initlst lsttotal, Array("C1", "C2", "C3"), Array(0.45, 0.4, 0.15)
  ws = Array(15, 10, 18, 25, 13, 19)
End Sub

Private Sub txtemp_finbusqueda(llave As String, valor As String)
  txtemp = llave
  labnom = valor
  crearegresos llave
  txtfecha(0).enabled = True: txtfecha(1).enabled = True
  llst
End Sub

Private Sub txtemp_vacio()
  labnom = ""
  txtfecha(0).enabled = False: txtfecha(1).enabled = False
  txtfecha(0) = "": txtfecha(1) = ""
  lstbrutos.ListItems.Clear
End Sub

Private Sub txtfecha_GotFocus(Index As Integer)
  txtfecha(Index).SelStart = 0
  txtfecha(Index).SelLength = 10
End Sub

Private Sub Command2_Click()
  Dim i As Integer, j As Integer, k As Integer, t As String
  On Error GoTo E
  selimpr.Show vbModal
  If Not selimpr.Cancel Then
    k = 0: titulo k: n = 0
    For i = 1 To lstbrutos.ListItems.Count
      t = left2(lstbrutos.ListItems(i), ws(0)) & " "
      For j = 1 To lstbrutos.ListItems(i).ListSubItems.Count
        t = t & IIf(j >= 5, right2(Format(lstbrutos.ListItems(i).ListSubItems(j), "0.00"), ws(j)), _
                            left2(lstbrutos.ListItems(i).ListSubItems(j), ws(j))) & " "
      Next
      n = lstbrutos.ListItems(i).ListSubItems(5)
      yx i + 7, 4, t
      If i > Printer.ScaleHeight - 3 Then
        parciales Printer.ScaleHeight - 3
        Printer.NewPage
        k = k + 1
        titulo k
        parciales 4
      End If
    Next
    parciales Printer.ScaleHeight - 3
    Printer.EndDoc
  End If
  Exit Sub
E: MsgBox "Error en la impresión: " & Err.Description, vbCritical, ""
End Sub

Private Sub titulo(ByVal p As Integer)
  Dim i As Integer, t As String, co As ColumnHeader
  yx 1, 4, "HOJA " & (p + 1)
  centro "PERCEPCIONES DE INGRESOS BRUTOS SOBRE COMPRAS"
  yx 2, 0, "": centro UCase(labnom)
  If txtfecha(0) <> "  /  /    " Then t = t & " DESDE EL " & txtfecha(0)
  If txtfecha(1) <> "  /  /    " Then t = t & " HASTA EL " & txtfecha(1)
  derecha t
  parciales 3
  For i = 1 To lstbrutos.ColumnHeaders.Count
    Set co = lstbrutos.ColumnHeaders(i)
    t = t & IIf(i >= 6, right2(co, ws(i - 1)), left2(co, ws(i - 1))) & " "
  Next
  yx 6, 4, t
  Printer.Line (4, 7)-(Printer.ScaleWidth - 4, 7)
End Sub

Private Sub parciales(ByVal l As Integer)
  Dim t As String
  Printer.Line (4, l)-(Printer.ScaleWidth - 4, l)
  t = String(ws(3) + ws(2) + ws(1) + ws(0) - 9, " ") & "    PARCIALES" & String(ws(4), " ") & " "
  t = t & right2(Format(n, "0.00"), ws(5)) & " "
  yx l + 1, 4, t
  Printer.Line (4, l + 2)-(Printer.ScaleWidth - 4, l + 2)
End Sub

Private Sub llst()
  Dim sql As String, i As ListItem
  sql = "select cod_egr,fecha,nom_comp,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero,nom_prov,format(cuit_prov,'00-00000000-0') as cuit_prov,perc_ib " & _
    "from ((egresos" & txtemp & " as e inner join proveedores as p on e.cod_prov=p.cod_prov) " & _
    "inner join comprobantes as c on c.cod_comp=e.letra) where perc_ib>0"
  If txtfecha(0) <> "  /  /    " Then sql = sql & " and fecha>=#" & Format(txtfecha(0), "mm/dd/yyyy") & "#"
  If txtfecha(1) <> "  /  /    " Then sql = sql & " and fecha<=#" & Format(txtfecha(1), "mm/dd/yyyy") & "#"
  sql = sql & " order by fecha asc"
  llenarlst lstbrutos, sql, Array("fecha", "nom_comp", "numero", "nom_prov", "cuit_prov", "perc_ib"), "cod_egr"
  n = 0: For Each i In lstbrutos.ListItems: n = n + i.ListSubItems(5): Next
  lsttotal.ListItems.Clear
  With lsttotal.ListItems.Add
    .ListSubItems.Add , , "TOTAL DE PERCEPCIONES"
    .ListSubItems.Add , , n
  End With
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
  llst
End Sub
