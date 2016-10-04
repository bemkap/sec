VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "Editar"
      Enabled         =   0   'False
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
      Left            =   3243
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdgraficar 
      Caption         =   "Graficar"
      Enabled         =   0   'False
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
      Left            =   4560
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
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
      Height          =   1575
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
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
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtperiodo 
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
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "per."
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
         TabIndex        =   1
         Top             =   945
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "año"
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
         Top             =   480
         Width           =   675
      End
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
   Begin Project1.UserControl1 txtemp 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      info            =   "Ingresar código de empresa. F3: buscar"
      tabla           =   "empresas"
      campo           =   "nom_emp"
      clave           =   "cod_emp"
      busq            =   "nom_emp"
      regvalid        =   "regvalid"
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
      Left            =   3600
      TabIndex        =   10
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
      Left            =   840
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "calculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbperiodos_Click()
  If cmbperiodos <> "" Then llenar cmbperiodos
End Sub

Private Sub cmdgraficar_Click()
  On Error GoTo E
  assert txtemp <> "" And txtperiodo <> "", NOCAMP, "Campos obligatorios: empresa y año"
  fhisto.aa = txtperiodo
  fhisto.Show vbModal
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub cmdnuevo_Click()
  On Error GoTo E
  assert txtperiodo <> "" And cmbperiodos.ListIndex > -1, NOCAMP, "Campos obligatorios: año y mes"
  initlst selcierre.lstcomp(0), Array("FECHA", "Nº COMPROBANTE", "IVA", "SUBTOTAL"), _
    Array(0.25, 0.25, 0.25, 0.25)
  initlst selcierre.lstcomp(1), Array("FECHA", "Nº COMPROBANTE", "IVA", "SUBTOTAL"), _
    Array(0.25, 0.25, 0.25, 0.25)
  initlst selcierre.lstcomp1(0), Array("FECHA", "Nº COMPROBANTE", "IVA", "SUBTOTAL"), _
    Array(0.25, 0.25, 0.25, 0.25)
  initlst selcierre.lstcomp1(1), Array("FECHA", "Nº COMPROBANTE", "IVA", "SUBTOTAL"), _
    Array(0.25, 0.25, 0.25, 0.25)
  
  llenarlst selcierre.lstcomp(0), "select cod_egr,fecha,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero," & _
    "iva21+iva105+iva27 as iva,gravado+no_gravado+iva21+iva105+iva27+exento+interno+perc_iva+perc_ib" & _
    "+litros*0.27 as subtotal from egresos" & txtemp & " where periodo=" & txtperiodo * 12 + cmbperiodos, _
    Array("fecha", "numero", "iva", "subtotal"), "cod_egr"
  llenarlst selcierre.lstcomp1(0), "select cod_egr,fecha,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero," & _
    "iva21+iva105+iva27 as iva,gravado+no_gravado+iva21+iva105+iva27+exento+interno+perc_iva+perc_ib" & _
    "+litros*0.27 as subtotal from egresos" & txtemp & " where periodo<=0", _
    Array("fecha", "numero", "iva", "subtotal"), "cod_egr"
  llenarlst selcierre.lstcomp(1), "select cod_ing,fecha,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero," & _
    "iva21+iva105+iva27 as iva,gravado+no_gravado+iva21+iva105+iva27+exento+interno+ret_iva+ret_ib" & _
    " as subtotal from ingresos" & txtemp & " where periodo=" & txtperiodo * 12 + cmbperiodos, _
    Array("fecha", "numero", "iva", "subtotal"), "cod_ing"
  llenarlst selcierre.lstcomp1(1), "select cod_ing,fecha,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero," & _
    "iva21+iva105+iva27 as iva,gravado+no_gravado+iva21+iva105+iva27+exento+interno+ret_iva+ret_ib" & _
    " as subtotal from ingresos" & txtemp & " where periodo<=0", _
    Array("fecha", "numero", "iva", "subtotal"), "cod_ing"
  selcierre.periodo = txtperiodo * 12 + cmbperiodos
  selcierre.emp = txtemp
  selcierre.Show vbModal
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  Dim col(), fil(), i As Integer
  col = Array("COMPRAS", "VENTAS")
  fil = Array("GRAVADO", "NO GRAVADO", "IVA", "EXENTO", "INTERNO", "PERC./RET. IVA", "PERC./RET. IB", "LITROS*0.27", "SUBTOTAL")
  With flx1
    For i = 0 To UBound(fil): .TextMatrix(i + 1, 0) = fil(i): Next
    For i = 0 To UBound(col): .TextMatrix(0, i + 1) = col(i): Next
    .ColWidth(0) = 2000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
  End With
End Sub

Private Sub txtemp_vacio()
  labnom = ""
  txtperiodo.enabled = False
  cmdgraficar.enabled = False
  cmdnuevo.enabled = False
End Sub

Private Sub txtperiodo_Change()
  cmbperiodos.Clear
  cmbperiodos.enabled = False
End Sub

Private Sub txtperiodo_GotFocus()
  txtperiodo.SelStart = 0
  txtperiodo.SelLength = Len(txtperiodo.text)
End Sub

Private Sub txtemp_finbusqueda(llave As String, valor As String)
  txtemp = llave
  labnom = valor
  txtperiodo.enabled = True
  txtperiodo.SetFocus
  cmdgraficar.enabled = True
  cmdnuevo.enabled = True
End Sub

Private Sub llenar(ByVal j As Integer)
  Dim n As Double, m As Double, i As Integer
  n = 0: m = 0
  With query("vte", , "periodo=" & txtperiodo * 12 + cmbperiodos)
    For i = 1 To flx1.Rows - 2
      flx1.TextMatrix(i, 1) = "0.00"
      If .RecordCount > 0 Then flx1.TextMatrix(i, 1) = Format(coalesce(.fields(i - 1), 0), "0.00")
      n = n + val(flx1.TextMatrix(i, 1))
    Next
  End With
  With query("vti", , "periodo=" & txtperiodo * 12 + cmbperiodos)
    For i = 1 To flx1.Rows - 2
      flx1.TextMatrix(i, 2) = "0.00"
      If .RecordCount > 0 Then flx1.TextMatrix(i, 2) = Format(coalesce(.fields(i - 1), 0), "0.00")
      m = m + val(flx1.TextMatrix(i, 2))
    Next
  End With
  flx1.TextMatrix(flx1.Rows - 1, 1) = n
  flx1.TextMatrix(flx1.Rows - 1, 2) = m
End Sub

Private Sub txtperiodo_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer
  On Error GoTo E
  If KeyCode = vbKeyReturn Then
    If txtperiodo <> "" Then
      For i = 1 To 12: cmbperiodos.AddItem i: Next
      If viewexiste("vti") Then C.Execute "drop view vti"
      If viewexiste("vte") Then C.Execute "drop view vte"
      C.Execute "create view vti as " & _
          "select sum(gravado) as sgravado,sum(no_gravado) as sno_gravado,sum(iva21)+sum(iva105)+sum(iva27) as siva," & _
          "sum(exento) as sexento,sum(interno) as sinterno,sum(ret_iva) as sret_iva,sum(ret_ib) as sret_ib,'' as slitros,periodo," & _
          "sum(iva21) as s21,sum(iva105) as s105,sum(iva27) as s27 from ingresos" & txtemp & _
          " where periodo>" & txtperiodo * 12 & " and periodo<=" & txtperiodo * 12 + 12 & " group by periodo"
      C.Execute "create view vte as " & _
          "select sum(gravado) as sgravado,sum(no_gravado) as sno_gravado,sum(iva21)+sum(iva105)+sum(iva27) as siva," & _
          "sum(exento) as sexento,sum(interno) as sinterno,sum(perc_iva) as sperc_iva,sum(perc_ib) as sperc_ib,sum(litros)*0.27 as slitros,periodo," & _
          "sum(iva21) as s21,sum(iva105) as s105,sum(iva27) as s27 from egresos" & txtemp & _
          " where periodo>" & txtperiodo * 12 & " and periodo<=" & txtperiodo * 12 + 12 & " group by periodo"
      cmbperiodos.enabled = (txtperiodo <> "")
      cmbperiodos.ListIndex = -1
      cmbperiodos.SetFocus
    End If
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub
