VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form duplicidad 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
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
      Left            =   1882
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "COMPRAS - GASTOS"
      TabPicture(0)   =   "duplicidad.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(3)=   "txtfecha(1)"
      Tab(0).Control(4)=   "txtfecha(0)"
      Tab(0).Control(5)=   "lstcompras"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "VENTAS - COBROS"
      TabPicture(1)   =   "duplicidad.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtfecha(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtfecha(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lstventas"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin MSComctlLib.ListView lstcompras 
         Height          =   4935
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
      Begin MSComctlLib.ListView lstventas 
         Height          =   4935
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
         Index           =   2
         Left            =   1920
         TabIndex        =   3
         Top             =   5520
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
         Index           =   3
         Left            =   4080
         TabIndex        =   4
         Top             =   5520
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
         Index           =   0
         Left            =   -73080
         TabIndex        =   1
         Top             =   5520
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
         Left            =   -70920
         TabIndex        =   2
         Top             =   5520
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
      Begin VB.Label Label6 
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
         Left            =   -71640
         TabIndex        =   13
         Top             =   5640
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   -74760
         TabIndex        =   12
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Left            =   -73920
         TabIndex        =   11
         Top             =   5640
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
         Left            =   3360
         TabIndex        =   10
         Top             =   5640
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
         Left            =   240
         TabIndex        =   9
         Top             =   5640
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
         TabIndex        =   8
         Top             =   5640
         Width           =   735
      End
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
      Left            =   3562
      TabIndex        =   15
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label7 
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
      Left            =   727
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "duplicidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  initlst lstcompras, Array("FECHA", "TIPO", "Nº", "PROVEEDOR", "CUIT", "TOTAL"), _
    Array(0.13, 0.1, 0.18, 0.25, 0.17, 0.15)
  initlst lstventas, Array("FECHA", "TIPO", "Nº", "CLIENTE", "CUIT", "TOTAL"), _
    Array(0.13, 0.1, 0.18, 0.25, 0.17, 0.15)
End Sub

Private Sub txtemp_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de empresa. F3: buscar"
End Sub

Private Sub txtemp_KeyDown(keycode As Integer, Shift As Integer)
  If keycode = vbKeyF3 Then
    If teclaemp(txtemp, labnom) Then cargar
  End If
End Sub

Private Sub txtemp_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtemp_Validate(Cancel As Boolean)
  If txtemp <> "" Then Cancel = validaremp(txtemp, labnom) Else labnom = ""
  If Not Cancel And txtemp <> "" Then cargar
End Sub

Private Sub txtfecha_GotFocus(index As Integer)
  txtfecha(index).SelStart = 0
  txtfecha(index).SelLength = 10
End Sub

Private Sub txtfecha_LostFocus(index As Integer)
  Select Case index
  Case 0, 1: fecha1
  Case 2, 3: fecha2
  End Select
End Sub

Private Sub fecha1()
  'filtra desde fecha1
  If viewexiste("vegresos") Then
    sql = "select * from vegresos where true"
    If txtfecha(0) <> "  /  /    " Then sql = sql & " and fecha>=#" & Format(txtfecha(0), "mm/dd/yyyy") & "#"
    If txtfecha(1) <> "  /  /    " Then sql = sql & " and fecha<=#" & Format(txtfecha(1), "mm/dd/yyyy") & "#"
    llenarlst lstcompras, CStr(sql), Array("fecha", "nom_comp", "numero", "nom_prov", "cuit_prov1", "subtotal"), "cod_egr"
  End If
End Sub

Private Sub fecha2()
  'filtra hasta fecha2
  If viewexiste("vingresos") Then
    sql = "select * from vingresos where true"
    If txtfecha(2) <> "  /  /    " Then sql = sql & " and fecha>=#" & Format(txtfecha(2), "mm/dd/yyyy") & "#"
    If txtfecha(3) <> "  /  /    " Then sql = sql & " and fecha<=#" & Format(txtfecha(3), "mm/dd/yyyy") & "#"
    llenarlst lstventas, CStr(sql), Array("fecha", "nom_comp", "numero", "nom_cli", "cuit_cli1", "subtotal"), "cod_ing"
  End If
End Sub

Private Sub cargar()
  'crea tablas de consulta para los ingresos y egresos duplicados
  crearingresos txtemp: crearegresos txtemp
  If tablaexiste("degresos" & txtemp) Then
    If viewexiste("vegresos") Then C.Execute "drop view vegresos"
    C.Execute "create view vegresos as " & _
      "select cod_egr,fecha,nom_comp,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero," & _
      "nom_prov,format(cuit_prov,'00-00000000-0') as cuit_prov1," & _
      "format(gravado+no_gravado+exento+interno+litros*0.27+iva21+iva105+iva27+perc_iva+perc_ib,'0.00') as subtotal " & _
      "from ((degresos" & txtemp & " as e " & _
      "inner join comprobantes as c on e.letra=c.cod_comp) " & _
      "inner join proveedores as p on e.cod_prov=p.cod_prov)"
    llenarlst lstcompras, "select * from vegresos", Array("fecha", "nom_comp", "numero", "nom_prov", "cuit_prov1", "subtotal"), "cod_egr"
  End If
  If tablaexiste("dingresos" & txtemp) Then
    If viewexiste("vingresos") Then C.Execute "drop view vingresos"
    C.Execute "create view vingresos as " & _
      "select cod_ing,fecha,nom_comp,format(sucursal,'0000')&'-'&format(n_comp,'00000000') as numero," & _
      "nom_cli,format(cuit_cli,'00-00000000-0') as cuit_cli1," & _
      "format(gravado+no_gravado+exento+interno+iva21+iva105+iva27+ret_iva+ret_ib,'0.00') as subtotal " & _
      "from ((dingresos" & txtemp & " as i " & _
      "inner join comprobantes as c on i.letra=c.cod_comp) " & _
      "inner join clientes as p on i.cod_cli=p.cod_cli)"
    llenarlst lstventas, "select * from vingresos", Array("fecha", "nom_comp", "numero", "nom_cli", "cuit_cli1", "subtotal"), "cod_ing"
  End If
End Sub

Private Sub txtfecha_Validate(index As Integer, Cancel As Boolean)
  If txtfecha(index) <> "  /  /    " Then Cancel = Not validarfecha(txtfecha(index))
End Sub
