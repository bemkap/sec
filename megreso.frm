VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form megreso 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   60
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtcomp 
      Alignment       =   1  'Right Justify
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
      Left            =   2160
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtsucursal 
      Alignment       =   1  'Right Justify
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
      Left            =   2160
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ComboBox cmbletra 
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
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtproveedor 
      Alignment       =   1  'Right Justify
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
      Left            =   2160
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
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
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Index           =   7
      Left            =   5880
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtcuenta 
      Alignment       =   1  'Right Justify
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
      Left            =   5880
      TabIndex        =   15
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtsubtotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
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
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar"
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
      Left            =   4320
      TabIndex        =   17
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
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
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar"
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
      Left            =   3000
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtfecha 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Litros"
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
      TabIndex        =   40
      Top             =   3480
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
      TabIndex        =   39
      Top             =   2040
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
      TabIndex        =   38
      Top             =   1680
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
      TabIndex        =   37
      Top             =   3120
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
      TabIndex        =   36
      Top             =   2400
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
      TabIndex        =   35
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Perc. IVA"
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
      TabIndex        =   34
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Perc. IB."
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
      TabIndex        =   33
      Top             =   4200
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
      TabIndex        =   32
      Top             =   4560
      Width           =   1335
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
      TabIndex        =   31
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label17 
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
      TabIndex        =   30
      Top             =   600
      Width           =   975
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
      TabIndex        =   29
      Top             =   3120
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
      TabIndex        =   28
      Top             =   3480
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
      TabIndex        =   27
      Top             =   2760
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
      TabIndex        =   26
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Cód.proveedor"
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
      TabIndex        =   25
      Top             =   3840
      Width           =   1695
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
      Left            =   3480
      TabIndex        =   24
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label labiva 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DataSource      =   "Adodc1"
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
      TabIndex        =   23
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lablitros 
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
      Left            =   7800
      TabIndex        =   22
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label labproveedor 
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
      Left            =   2160
      TabIndex        =   21
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label labcodcue 
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
      Left            =   5880
      TabIndex        =   20
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Cód.compra"
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
      TabIndex        =   18
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "megreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private porcs(2, 1) As Double, idc As Integer, adoegr As ADODB.Recordset

Private Sub cmdguardar_Click()
  On Error GoTo E
  assert txtcodigo <> "" And txtemp <> "" And txtsucursal <> "" And txtcomp <> "" And txtfecha <> "  /  /    " And cmbletra.ListIndex <> -1, NOCAMP, "Campos obligatorios: código, empresa, sucursal, número, fecha y tipo"
  assert Not adoegr Is Nothing, NOCAMP, "Elegir comprobante para editar"
  guardaregreso Me, adoegr, porcs
  adoegr.Update
  StatusBar1.SimpleText = "Comprobante modificado"
  limpiaregreso Me
  idc = 0: porcs(0, 1) = 0: porcs(1, 1) = 0: porcs(2, 1) = 0
  labiva = porcs(0, 0) * 100 & "%"
  Set adoegr = Nothing
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  loadmov Me, porcs
End Sub

Private Sub txtcodigo_GotFocus()
  StatusBar1.SimpleText = "F3: buscar"
End Sub

Private Sub txtcodigo_KeyDown(keycode As Integer, Shift As Integer)
  On Error GoTo E
  assert txtemp <> "", NOCAMP, "Falta ingresar empresa"
  If keycode = vbKeyF3 Then
    formbuscar "egresos" & txtemp, "cod_egr|fecha|format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero", "cod_egr", "n_comp"
    If Not buscar.Cancel Then
      Set adoegr = busc("select * from egresos" & txtemp & " where cod_egr=" & buscar.val)
      With adoegr
        txtcodigo = buscar.val
        txtfecha = !fecha
        txtsucursal = !sucursal
        txtcomp = !n_comp
        txtproveedor = !cod_prov
        labproveedor = busc("select nom_prov from proveedores where cod_prov=" & !cod_prov)!nom_prov
        cmbletra.ListIndex = val(!letra)
        txtn(0) = Format(!gravado, "0.00")
        txtn(1) = Format(!no_gravado, "0.00")
        txtn(2) = Format(!exento, "0.00")
        txtn(3) = Format(!interno, "0.00")
        porcs(0, 1) = Format(!iva21, "0.00")
        porcs(1, 1) = Format(!iva105, "0.00")
        porcs(2, 1) = Format(!iva27, "0.00")
        txtn(4) = Format(porcs(idc, 1), "0.00")
        txtlitros = !litros
        txtn(6) = Format(!perc_iva, "0.00")
        txtn(7) = Format(!perc_ib, "0.00")
        txtcuenta = !cod_cue
        labcodcue = busc("select nom_cue from cuentas where cod_cue=" & !cod_cue)!nom_cue
      End With
      calcsubt
    End If
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub txtemp_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtemp_Validate(Cancel As Boolean)
  If txtemp <> "" Then Cancel = validaremp(txtemp, labnom) Else labnom = ""
  If Not Cancel And txtemp <> "" Then crearegresos txtemp
End Sub

Private Sub txtfecha_GotFocus()
  txtfecha.SelStart = 0
  txtfecha.SelLength = 10
End Sub

Private Sub txtfecha_Validate(Cancel As Boolean)
  If txtfecha <> "  /  /    " Then Cancel = Not validarfecha(txtfecha)
End Sub

Private Sub txtn_GotFocus(index As Integer)
  Select Case index
  Case 4:
    StatusBar1.SimpleText = "F2: cambiar alícuota"
  End Select
  txtn(index).SelStart = 0
  txtn(index).SelLength = Len(txtn(index))
End Sub

Private Sub txtn_KeyDown(index As Integer, keycode As Integer, Shift As Integer)
  If index = 4 And keycode = vbKeyF2 Then
    teclatxtnmov Me, 4, idc, porcs
    calcsubt
  End If
End Sub

Private Sub txtn_LostFocus(index As Integer)
  lfocustxtnegreso Me, index, idc, porcs
  calcsubt
End Sub

Private Sub txtProveedor_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de proveedor. F3: buscar. F4: agregar"
End Sub

Private Sub txtproveedor_KeyDown(keycode As Integer, Shift As Integer)
  Select Case keycode
  Case vbKeyF3: teclaprov txtproveedor, labproveedor
  Case vbKeyF4: If p And 2 ^ 2 Then teclaprov1 Else StatusBar1.SimpleText = "Permisos necesarios"
  End Select
End Sub

Private Sub txtproveedor_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtemp_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de empresa. F3: buscar"
End Sub

Private Sub txtemp_KeyDown(keycode As Integer, Shift As Integer)
  If keycode = vbKeyF3 Then
    If teclaemp(txtemp, labnom) Then crearegresos txtemp
  End If
End Sub

Private Sub txtcuenta_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de cuenta. F3: buscar. F4: agregar"
End Sub

Private Sub txtcuenta_KeyDown(keycode As Integer, Shift As Integer)
  teclacuemov Me, keycode
End Sub

Private Sub txtcuenta_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtcuenta_Validate(Cancel As Boolean)
  Cancel = validarcuemov(Me)
End Sub

Private Sub txtproveedor_Validate(Cancel As Boolean)
  If txtproveedor <> "" Then Cancel = validarprov(txtproveedor, labproveedor)
End Sub

Private Sub calcsubt()
  txtsubtotal = 0
  For i = 0 To txtn.UBound
    txtsubtotal = val(txtsubtotal) + IIf(i = 4 Or i = 5, 0, val(txtn(i)))
    txtn(i) = Format(val(txtn(i)), "0.00")
  Next
  txtsubtotal = val(txtsubtotal) + porcs(0, 1) + porcs(1, 1) + porcs(2, 1)
  txtsubtotal = Format(val(txtsubtotal) + val(txtn(5)) * 0.27, "0.00")
End Sub

