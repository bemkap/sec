VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form calculo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
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
      Left            =   1905
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Graficar"
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
      Left            =   4545
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ComboBox cmbperiodos 
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
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERIODO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
      Begin MSMask.MaskEdBox txtperiodo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtperiodo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   675
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
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Calcular"
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
      Left            =   3225
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   3735
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      RowHeightMin    =   350
      BackColorBkg    =   -2147483632
      GridLinesFixed  =   1
      AllowUserResizing=   1
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
      Left            =   3585
      TabIndex        =   11
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label3 
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
      Left            =   705
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "calculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tc As New Collection, tv As New Collection

Private Sub cmbperiodos_Click()
  llenar cmbperiodos.ListIndex + 1
End Sub

Private Sub cmdguardar_Click()
  On Error GoTo E
  assert txtemp <> "" And txtperiodo(0) <> "  /    " And txtperiodo(1) <> "  /    ", NOCAMP, "Campos obligatorios: empresa y periodos"
  assert CDate(txtperiodo(0)) <= CDate(txtperiodo(1)), INVDAT, "Fechas incorrectas"
  cmbperiodos.Enabled = True
  For m = Month(txtperiodo(0)) - 1 To Month(txtperiodo(1)) + 12 * (Year(txtperiodo(1)) - Year(txtperiodo(0))) - 1
    cmbperiodos.AddItem (m Mod 12) + 1 & "/" & Year(txtperiodo(0)) + Round((m + 1) / 12 - 0.51)
  Next
  For i = 0 To cmbperiodos.ListCount - 1
    mm = Month(CDate(cmbperiodos.List(i)))
    yy = Year(CDate(cmbperiodos.List(i)))
    'totales de compras por periodo
    tc.Add busc("select sum(gravado),sum(no_gravado),sum(iva21)+sum(iva105)+sum(iva27),sum(exento),sum(interno),sum(perc_iva),sum(perc_ib),sum(litros*0.27) from egresos" & txtemp & _
                " having month(fecha)=" & mm & " and year(fecha)=" & yy)
    'totales de ventas por periodo
    tv.Add busc("select sum(gravado),sum(no_gravado),sum(iva21)+sum(iva105)+sum(iva27),sum(exento),sum(interno),sum(ret_iva),sum(ret_ib) from ingresos" & txtemp & _
                " having month(fecha)=" & mm & " and year(fecha)=" & yy)
  Next
  cmbperiodos.ListIndex = 0
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Command1_Click()
  On Error GoTo E
  assert txtemp <> "" And txtperiodo(0) <> "  /    " And txtperiodo(1) <> "  /    ", NOCAMP, "Campos obligatorios: empresa y periodos"
  assert CDate(txtperiodo(0)) <= CDate(txtperiodo(1)), INVDAT, "Fechas incorrectas"
  'construccion de las cadenas para el histrograma
  For j = 1 To cmbperiodos.ListCount
    n = 0: m = 0
    For i = 0 To 7
      n = n + coalesce(tc(j).Fields(i), 0)
      If i < 7 Then m = m + coalesce(tv(j).Fields(i), 0)
    Next
    se = se & "," & n 'importe egreso(compra/gasto)
    si = si & "," & m 'importe ingreso(venta/cobro)
    my = my & "," & cmbperiodos.List(j - 1) 'mes
  Next
  fhisto.se = Mid(se, 2): fhisto.si = Mid(si, 2)
  fhisto.em = Mid(my, 2): fhisto.im = Mid(my, 2)
  fhisto.Show vbModal
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  col = Array("COMPRAS", "VENTAS")
  fil = Array("GRAVADO", "NO GRAVADO", "IVA", "EXENTO", "INTERNO", "PERC/RET IVA", "PERC./RET. IB", "LITROS*0.27", "SUBTOTAL")
  With flx1
    For i = 0 To UBound(fil): .TextMatrix(i + 1, 0) = fil(i): Next
    For i = 0 To UBound(col): .TextMatrix(0, i + 1) = col(i): Next
    .ColWidth(0) = 2000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
  End With
End Sub

Private Sub txtemp_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de empresa. F3: buscar"
End Sub

Private Sub txtemp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    If teclaemp(txtemp, labnom) Then
      For j = 0 To min(tc.Count, tv.Count) - 1: llenar j: Next
    End If
  End If
End Sub

Private Sub txtemp_Validate(Cancel As Boolean)
  If txtemp <> "" Then Cancel = validaremp(txtemp, labnom) Else labnom = ""
  If Not Cancel And txtemp <> "" Then
    For j = 0 To min(tc.Count, tv.Count) - 1: llenar j: Next
  End If
End Sub

Private Sub txtemp_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtperiodo_Change(Index As Integer)
  cmbperiodos.Clear
  cmbperiodos.Enabled = False
  If Not tv Is Nothing Then Set tv = Nothing: Set tv = New Collection
  If Not tc Is Nothing Then Set tc = Nothing: Set tc = New Collection
End Sub

Private Sub txtperiodo_GotFocus(Index As Integer)
  txtperiodo(Index).SelStart = 0
  txtperiodo(Index).SelLength = 7
End Sub

Private Sub llenar(ByVal j As Integer)
  n = 0: m = 0
  With flx1
    For i = 1 To .Rows - 2
      .TextMatrix(i, 1) = Format(coalesce(tc(j).Fields(i - 1), 0), "0.00")
      If i < .Rows - 2 Then .TextMatrix(i, 2) = Format(coalesce(tv(j).Fields(i - 1), 0), "0.00")
      n = n + val(.TextMatrix(i, 1))
      m = m + val(.TextMatrix(i, 2))
    Next
    .TextMatrix(.Rows - 1, 1) = n
    .TextMatrix(.Rows - 1, 2) = m
  End With
End Sub

Private Sub txtperiodo_Validate(Index As Integer, Cancel As Boolean)
  If txtperiodo(Index) <> "  /    " Then Cancel = Not validarfecha(txtperiodo(Index))
End Sub
