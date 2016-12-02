VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form impexp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Importar"
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
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   4560
      TabIndex        =   21
      Top             =   1200
      Width           =   4215
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "..."
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
         Left            =   3600
         TabIndex        =   22
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txttablai 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         IMEMode         =   3  'DISABLE
         Left            =   1200
         TabIndex        =   7
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton cmbimportar 
         Caption         =   "Importar"
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
         Left            =   1560
         TabIndex        =   23
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtarchivoi 
         Appearance      =   0  'Flat
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Tabla"
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
         TabIndex        =   25
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Archivo"
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
         TabIndex        =   24
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Exportar"
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
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   4215
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   1200
         TabIndex        =   27
         Top             =   3480
         Width           =   2775
         Begin VB.CheckBox chkborrar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Borrar registros"
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
            Left            =   30
            TabIndex        =   5
            Top             =   120
            Width           =   2715
         End
      End
      Begin VB.ComboBox cmbfmt 
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
         ItemData        =   "expimp.frx":0000
         Left            =   1200
         List            =   "expimp.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1110
         Left            =   1200
         TabIndex        =   2
         Top             =   1680
         Width           =   2775
         Begin VB.CheckBox chktabla 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ventas-Cobros(dup.)"
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
            Index           =   3
            Left            =   30
            TabIndex        =   20
            Top             =   840
            Width           =   2715
         End
         Begin VB.CheckBox chktabla 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ventas-Cobros"
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
            Index           =   2
            Left            =   30
            TabIndex        =   19
            Top             =   600
            Width           =   2715
         End
         Begin VB.CheckBox chktabla 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Compras-Gastos(dup.)"
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
            Index           =   1
            Left            =   30
            TabIndex        =   18
            Top             =   360
            Width           =   2715
         End
         Begin VB.CheckBox chktabla 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Compras-Gastos"
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
            Index           =   0
            Left            =   30
            TabIndex        =   17
            Top             =   120
            Width           =   2715
         End
      End
      Begin VB.CommandButton cmdexportar 
         Caption         =   "Exportar"
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
         Left            =   1560
         TabIndex        =   12
         Top             =   4800
         Width           =   1215
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
         Left            =   1200
         TabIndex        =   3
         Top             =   3000
         Width           =   975
         _ExtentX        =   1720
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
         Left            =   3000
         TabIndex        =   4
         Top             =   3000
         Width           =   975
         _ExtentX        =   1720
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Formato"
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
         TabIndex        =   26
         Top             =   1305
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Periodo"
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
         TabIndex        =   16
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   15
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         Left            =   2160
         TabIndex        =   14
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tabla"
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
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   480
      Width           =   7455
      Begin Project1.UserControl1 txtemp 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   0
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
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
         Left            =   2760
         TabIndex        =   10
         Top             =   0
         Width           =   4695
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
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Abrir"
      Filter          =   "CSV Files (*.csv)|*.csv"
      FilterIndex     =   1
   End
End
Attribute VB_Name = "impexp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbimportar_Click()
  On Error GoTo E
  assert txtarchivoi <> "", NOCAMP, "Ingresar nombre de archivo"
  importar dialog.FileName, txttablai
  StatusBar1.SimpleText = "Tabla importada"
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub cmdexportar_Click()
  On Error GoTo E
  Dim tag As String, fecha As String, rec As ADODB.Recordset
  Dim tablas(), claves(), tipos(), i As Integer
  assert cmbfmt.ListIndex > -1, NOCAMP, "Elegir formato"
  fecha = "true"
  If txtfecha(0) <> "  /    " Then fecha = fecha & " and periodo>=" & Month(CDate(txtfecha(0))) + 12 * Year(CDate(txtfecha(0)))
  If txtfecha(1) <> "  /    " Then fecha = fecha & " and periodo<=" & Month(CDate(txtfecha(1))) + 12 * Year(CDate(txtfecha(1)))
  tablas = Array("egresos", "degresos", "ingresos", "dingresos")
  claves = Array("cod_egr", "cod_egr", "cod_ing", "cod_ing")
  fmts = Array(0, 0, 0, 0, 1, 1, 2, 2)
  For i = 0 To 3
    If chktabla(i) Then
      Set rec = query(tablas(i) & txtemp, , fecha, claves(i))
      If rec.RecordCount > 0 Then
        tag = rec.fields(claves(i)): rec.MoveLast
        tag = tag & "-" & rec.fields(claves(i)): rec.MoveFirst
        exportar rec, tablas(i) & txtemp & "-" & tag, CInt(fmts(cmbfmt.ListIndex * 4 + i))
        If chkborrar.Value = vbChecked Then C.Execute "delete from " & tablas(i) & txtemp & " where " & fecha
      End If
    End If
  Next
  StatusBar1.SimpleText = "Tablas exportadas"
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Command1_Click()
  On Error GoTo E
  dialog.ShowOpen
  txtarchivoi = dialog.FileTitle
  txttablai = left(txtarchivoi, InStr(1, txtarchivoi, "-") - 1)
E:
End Sub

Private Sub txtemp_finbusqueda(llave As String, valor As String)
  Frame1.enabled = True
  Frame4.enabled = True
  txtemp = llave
  labnom = valor
  cmbfmt.SetFocus
End Sub

Private Sub txtemp_vacio()
  Frame1.enabled = False
  Frame4.enabled = False
End Sub
