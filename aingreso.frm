VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form abmingreso 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkcalc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Calcular IVA"
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   38
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin Project1.UserControl3 txtcuenta 
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   4200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      info            =   "Ingresar código de cuenta. F3: buscar"
      enabled         =   0   'False
   End
   Begin Project1.UserControl1 txtemp 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   480
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
   Begin Project1.UserControl1 txtcodigo 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      info            =   "F3: buscar"
      campo           =   "cod_ing|fecha|format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero"
      clave           =   "cod_ing"
      busq            =   "n_comp"
   End
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "Eliminar"
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
      Left            =   4545
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Registrar"
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
      Left            =   3225
      TabIndex        =   15
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Index           =   3
      Left            =   5880
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   2760
      Width           =   1935
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtsucursal 
      Alignment       =   1  'Right Justify
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtcomp 
      Alignment       =   1  'Right Justify
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
      Left            =   2160
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Index           =   0
      Left            =   5880
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Index           =   2
      Left            =   5880
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Index           =   4
      Left            =   5880
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Index           =   5
      Left            =   5880
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Index           =   6
      Left            =   5880
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtsubtotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   5880
      TabIndex        =   18
      Text            =   "0.00"
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Index           =   1
      Left            =   5880
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   2040
      Width           =   1935
   End
   Begin Project1.UserControl1 txtcliente 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
      _extentx        =   1931
      _extenty        =   450
      info            =   "Ingresar código de cliente. F3: buscar. F4: agregar"
      tabla           =   "clientes"
      campo           =   "nom_cli"
      clave           =   "cod_cli"
      busq            =   "nom_cli"
      regvalid        =   "regvalid"
      enabled         =   0   'False
   End
   Begin MSMask.MaskEdBox txtfecha 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Label labcodigo 
      Alignment       =   1  'Right Justify
      Caption         =   "Cód.ingreso"
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
      TabIndex        =   17
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      TabIndex        =   37
      Top             =   600
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
      TabIndex        =   36
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label labcliente 
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
      Left            =   2160
      TabIndex        =   35
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label labcodcue 
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
      Left            =   5880
      TabIndex        =   34
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label labiva 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "21%"
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
      Left            =   7800
      TabIndex        =   33
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Cód.cliente"
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
      TabIndex        =   32
      Top             =   3600
      Width           =   1695
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
      Left            =   240
      TabIndex        =   31
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha"
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
      TabIndex        =   30
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Nº comprobante"
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
      TabIndex        =   29
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Sucursal"
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
      TabIndex        =   28
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Subtotal"
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
      Left            =   4320
      TabIndex        =   27
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Cód.cuenta"
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
      Left            =   4320
      TabIndex        =   26
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Ret. IB"
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
      Left            =   4320
      TabIndex        =   25
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Ret. IVA"
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
      Left            =   4320
      TabIndex        =   24
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Interno"
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
      Left            =   4320
      TabIndex        =   23
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Exento"
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
      Left            =   4320
      TabIndex        =   22
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "IVA"
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
      Left            =   4320
      TabIndex        =   21
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Gravado"
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
      Left            =   4320
      TabIndex        =   20
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "No gravado"
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
      Left            =   4320
      TabIndex        =   19
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "abmingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private porcs(2, 1) As Double, idc As Integer, ivadisc As Boolean, adoing As ADODB.Recordset
Public alta As Boolean

Private Sub cmbletra_Click()
  Dim i As Integer
  For i = 0 To txtn.UBound: txtn(i).enabled = (cmbletra.ListIndex > -1): Next
  txtcuenta.enabled = (cmbletra.ListIndex > -1)
  chkcalc.enabled = (cmbletra.ListIndex > -1)
  If cmbletra.ListIndex > -1 Then ivadisc = cmbletra.ItemData(cmbletra.ListIndex)
  For i = 0 To txtn.UBound: txtn(i) = "0.00": Next
  idc = 0: porcs(0, 1) = 0: porcs(1, 1) = 0: porcs(2, 1) = 0
  labiva = porcs(0, 0) * 100 & "%"
End Sub

Private Sub cmdeliminar_Click()
  On Error GoTo E
  assert Not adoing Is Nothing, NOCAMP, "Elegir comprobante a eliminar"
  If MsgBox("¿Realmente desea eliminar el comprobante?", vbYesNo) = vbYes Then
    adoing.Delete
    adoing.Update
    limpiaringreso
    idc = 0: porcs(0, 1) = 0: porcs(1, 1) = 0: porcs(2, 1) = 0
    labiva = porcs(0, 0) * 100 & "%"
    Set adoing = Nothing
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub cmdguardar_Click()
  Dim r As ADODB.Recordset, msg As String, tbl As String
  On Error GoTo E
  assert txtsucursal <> "" And txtcomp <> "" And txtfecha <> "  /  /    " And cmbletra.ListIndex <> -1, NOCAMP, "Campos obligatorios: empresa, sucursal, número, fecha y tipo"
  If alta Then
    'revisar duplicidad. el criterio para ser duplicado es tener la misma sucursal y numero
    Set r = query("ingresos" & txtemp, , "sucursal=" & txtsucursal & " and n_comp=" & txtcomp)
    msg = IIf(r.RecordCount > 0, "El comprobante ya existe. Se registra en la tabla duplicados", "Comprobante registrado")
    tbl = IIf(r.RecordCount > 0, "dingresos", "ingresos")
    Set adoing = tabl(tbl & txtemp)
    adoing.AddNew
    If r.RecordCount > 0 Then adoing!cod_ing = r!cod_ing
  Else
    assert Not adoing Is Nothing, NOCAMP, "Elegir comprobante para editar"
    msg = "Comprobante modificado"
  End If
  guardaringreso
  adoing.Update
  StatusBar1.SimpleText = msg
  limpiaringreso
  idc = 0: porcs(0, 1) = 0: porcs(1, 1) = 0: porcs(2, 1) = 0
  labiva = porcs(0, 0) * 100 & "%"
  Set adoing = Nothing
  If alta Then txtfecha.SetFocus Else txtcodigo.SetFocus
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  labcodigo.Visible = Not alta
  txtcodigo.Visible = Not alta
  cmdeliminar.Visible = Not alta
  porcs(0, 0) = 0.21: porcs(1, 0) = 0.105: porcs(2, 0) = 0.27
  llenarcmb cmbletra, "select * from comprobantes", "nom_comp", "ivadisc_comp"
End Sub

Private Sub txtcliente_finbusqueda(llave As String, valor As String)
  txtcliente = llave
  labcliente = left2(valor, 15)
  cmbletra.SetFocus
End Sub

Private Sub txtcliente_vacio()
  txtcliente = ""
  labcliente = ""
End Sub

Private Sub txtcodigo_finbusqueda(llave As String, valor As String)
  Set adoing = query("ingresos" & txtemp, , "cod_ing=" & valor)
  With adoing
    txtcodigo = valor
    txtfecha = !fecha
    txtsucursal = !sucursal
    txtcomp = !n_comp
    txtcliente = !cod_cli
    labcliente = query("clientes", "nom_cli", "cod_cli=" & !cod_cli)!nom_cli
    cmbletra.ListIndex = CDbl(!letra)
    txtn(0) = Format(!gravado, "0.00")
    txtn(1) = Format(!no_gravado, "0.00")
    txtn(2) = Format(!exento, "0.00")
    porcs(0, 1) = Format(!iva21, "0.00")
    porcs(1, 1) = Format(!iva105, "0.00")
    porcs(2, 1) = Format(!iva27, "0.00")
    txtn(3) = Format(porcs(idc, 1), "0.00")
    txtn(4) = Format(!interno, "0.00")
    txtn(5) = Format(!ret_iva, "0.00")
    txtn(6) = Format(!ret_ib, "0.00")
    If Not IsNull(!cod_cue) Then
      txtcuenta = !cod_cue
      labcodcue = query("cuentas", "nom_cue", "cod_cue=" & !cod_cue)!nom_cue
    End If
  End With
  calcsubt
End Sub

Private Sub txtcuenta_finbusqueda(llave As String, valor As String)
  txtcuenta = llave
  labcodcue = valor
  cmdguardar.SetFocus
End Sub

Private Sub txtcuenta_vacio()
  txtcuenta = ""
  labcodcue = ""
End Sub

Private Sub txtemp_finbusqueda(llave As String, valor As String)
  txtemp = llave
  labnom = valor
  crearingresos llave
  txtcodigo.tabla = "ingresos" & llave
  enable True
  txtcuenta.empresa = txtemp
  If alta Then txtfecha.SetFocus Else txtcodigo.SetFocus
End Sub

Private Sub txtemp_vacio()
  enable False
  labnom = ""
  limpiaringreso
End Sub

Private Sub txtfecha_GotFocus()
  txtfecha.SelStart = 0
  txtfecha.SelLength = 10
End Sub

Private Sub txtfecha_Validate(Cancel As Boolean)
  If txtfecha <> "  /  /    " Then Cancel = Not validarfecha(txtfecha)
End Sub

Private Sub txtn_GotFocus(Index As Integer)
  If Index = 3 Then StatusBar1.SimpleText = "F2: cambiar alícuota"
  txtn(Index).SelStart = 0
  txtn(Index).SelLength = Len(txtn(Index))
End Sub

Private Sub txtn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 3 And KeyCode = vbKeyF2 Then
    porcs(idc, 1) = CDbl(txtn(3))
    idc = (idc + 1) Mod 3
    txtn(3) = Format(porcs(idc, 1), "0.00")
    labiva = porcs(idc, 0) * 100 & "%"
    txtn(3).SelStart = 0
    txtn(3).SelLength = Len(txtn(3))
    calcsubt
  End If
End Sub

Private Sub txtn_LostFocus(Index As Integer)
  Select Case Index
  Case 0: 'gravado
    If chkcalc.Value = vbChecked Then
      If ivadisc Then
        porcs(0, 1) = CDbl(txtn(0)) * 0.21
      Else
        txtn(0).tag = txtn(0)
        txtn(0) = Format(CDbl(txtn(0)) / 1.21, "0.00")
        porcs(0, 1) = CDbl(txtn(0).tag) - CDbl(txtn(0))
      End If
      If idc = 0 Then txtn(3) = Format(porcs(0, 1), "0.00")
    End If
  Case 3: 'iva
    porcs(idc, 1) = CDbl(txtn(3))
    StatusBar1.SimpleText = ""
  End Select
  txtn(Index) = Format(txtn(Index), "0.00")
  calcsubt
End Sub

Private Sub txtcliente_alta()
  If p And 2 ^ 2 Then
    abmcliente.tmp = True
    abmcliente.alta = True
    abrir inicio.Frame1, abmcliente, False
  Else
    StatusBar1.SimpleText = "Permisos necesarios"
  End If
End Sub

Private Sub calcsubt()
  Dim i As Integer
  txtsubtotal = 0
  For i = 0 To txtn.UBound
    txtsubtotal = CDbl(txtsubtotal) + IIf(i = 3, 0, CDbl(txtn(i)))
    txtn(i) = Format(CDbl(txtn(i)), "0.00")
  Next
  txtsubtotal = CDbl(txtsubtotal) + porcs(0, 1) + porcs(1, 1) + porcs(2, 1)
  txtsubtotal = Format(txtsubtotal, "0.00")
End Sub

Private Sub enable(b As Boolean)
  txtcodigo.enabled = b
  txtfecha.enabled = b
  txtsucursal.enabled = b
  txtcomp.enabled = b
  txtcliente.enabled = b
  cmbletra.enabled = b
  cmdguardar.enabled = b
  cmdeliminar.enabled = b
  If Not b Then cmbletra.ListIndex = -1
End Sub

Public Sub limpiaringreso()
  Dim i As Integer
  txtcodigo = "": txtsucursal = "": txtcomp = "": txtfecha = "  /  /    "
  cmbletra.ListIndex = -1: txtcliente = ""
  labcliente = "": txtcuenta = "": labcodcue = ""
  txtsubtotal = "0.00"
  For i = 0 To txtn.UBound
    txtn(i) = "0.00"
    txtn(i).enabled = False
  Next
End Sub

Public Sub guardaringreso()
  With adoing
    !cod_emp = txtemp
    !sucursal = txtsucursal
    !n_comp = txtcomp
    !fecha = txtfecha
    !letra = cmbletra.ListIndex
    !cod_cli = txtcliente
    !gravado = CDbl(txtn(0))
    !no_gravado = CDbl(txtn(1))
    !exento = CDbl(txtn(2))
    !iva21 = porcs(0, 1)
    !iva105 = porcs(1, 1)
    !iva27 = porcs(2, 1)
    !interno = CDbl(txtn(4))
    !ret_iva = CDbl(txtn(5))
    !ret_ib = CDbl(txtn(6))
    If txtcuenta <> "" Then !cod_cue = CDbl(txtcuenta)
  End With
End Sub
