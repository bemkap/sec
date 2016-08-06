VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form aempresa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12303
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DATOS"
      TabPicture(0)   =   "aempresa.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtnom"
      Tab(0).Control(1)=   "txtcuit"
      Tab(0).Control(2)=   "txtloc"
      Tab(0).Control(3)=   "txtresp"
      Tab(0).Control(4)=   "txtcar"
      Tab(0).Control(5)=   "txtsus"
      Tab(0).Control(6)=   "txttel"
      Tab(0).Control(7)=   "txtdom"
      Tab(0).Control(8)=   "Label24"
      Tab(0).Control(9)=   "Label23"
      Tab(0).Control(10)=   "Label22"
      Tab(0).Control(11)=   "Label21"
      Tab(0).Control(12)=   "Label20"
      Tab(0).Control(13)=   "Label19"
      Tab(0).Control(14)=   "Label18"
      Tab(0).Control(15)=   "Label17"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "PLAN DE CUENTAS"
      TabPicture(1)   =   "aempresa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "trdisponibles"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ACTIVIDADES"
      TabPicture(2)   =   "aempresa.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "labactividad(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "labactividad(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "labactividad(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lstactividades"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Timer1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Command6"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtbuscar"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdguardar"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Picture1"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Picture2"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   8520
         Picture         =   "aempresa.frx":0054
         ScaleHeight     =   330
         ScaleWidth      =   345
         TabIndex        =   29
         Top             =   5040
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   11760
         ScaleHeight     =   330
         ScaleWidth      =   345
         TabIndex        =   28
         Top             =   6720
         Width           =   375
      End
      Begin VB.TextBox txtnom 
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
         Left            =   -70680
         TabIndex        =   1
         Top             =   2400
         Width           =   2895
      End
      Begin MSMask.MaskEdBox txtcuit 
         Height          =   375
         Left            =   -70680
         TabIndex        =   0
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-########-#"
         PromptChar      =   " "
      End
      Begin VB.CommandButton cmdguardar 
         Caption         =   "Registrar"
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
         Left            =   3990
         TabIndex        =   9
         Top             =   6480
         Width           =   1215
      End
      Begin VB.TextBox txtbuscar 
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
         Left            =   120
         TabIndex        =   19
         Top             =   5040
         Width           =   8775
      End
      Begin VB.CommandButton Command6 
         Height          =   255
         Left            =   7890
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5580
         Width           =   255
      End
      Begin VB.TextBox txtloc 
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
         Left            =   -70680
         TabIndex        =   3
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox txtresp 
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
         Left            =   -70680
         TabIndex        =   7
         Top             =   4560
         Width           =   2895
      End
      Begin VB.TextBox txtcar 
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
         Left            =   -70680
         TabIndex        =   6
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox txtsus 
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
         Left            =   -70680
         TabIndex        =   5
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox txttel 
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
         Left            =   -70680
         TabIndex        =   4
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox txtdom 
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
         Left            =   -70680
         TabIndex        =   2
         Top             =   2760
         Width           =   2895
      End
      Begin MSComctlLib.TreeView trdisponibles 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   11033
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   90
         LabelEdit       =   1
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
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   11760
         Top             =   120
      End
      Begin MSComctlLib.ListView lstactividades 
         Height          =   4575
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label labactividad 
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
         Height          =   255
         Index           =   2
         Left            =   2370
         TabIndex        =   26
         Top             =   6060
         Width           =   5775
      End
      Begin VB.Label labactividad 
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
         Height          =   255
         Index           =   1
         Left            =   2370
         TabIndex        =   25
         Top             =   5820
         Width           =   5775
      End
      Begin VB.Label labactividad 
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
         Height          =   255
         Index           =   0
         Left            =   2370
         TabIndex        =   24
         Top             =   5580
         Width           =   5775
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Terciaria"
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
         Height          =   255
         Left            =   1050
         TabIndex        =   23
         Top             =   6060
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Secundaria"
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
         Height          =   255
         Left            =   1050
         TabIndex        =   22
         Top             =   5820
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Primaria"
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
         Height          =   255
         Left            =   1050
         TabIndex        =   21
         Top             =   5580
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Responsabilidad IVA"
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
         Left            =   -73200
         TabIndex        =   18
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Carácter"
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
         Left            =   -73200
         TabIndex        =   17
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Suscribe"
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
         Left            =   -73200
         TabIndex        =   16
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Teléfono"
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
         Left            =   -73200
         TabIndex        =   15
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Localidad"
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
         Left            =   -73200
         TabIndex        =   14
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Domicilio"
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
         Left            =   -73200
         TabIndex        =   13
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Razón social"
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
         Left            =   -73200
         TabIndex        =   12
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "CUIT"
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
         Left            =   -73200
         TabIndex        =   11
         Top             =   2160
         Width           =   2295
      End
   End
End
Attribute VB_Name = "aempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim act As Integer

Private Sub cmdguardar_Click()
  On Error GoTo E
  Dim co As Integer, n As Node
  With tabl("empresas")
    .AddNew Array("cuit_emp", "nom_emp", "dom_emp", "loc_emp", "tel_emp", "sus_emp", "car_emp", "resp_emp"), _
            Array(txtcuit, txtnom, txtdom, txtloc, txttel, txtsus, txtcar, txtresp)
    .Update
    co = !cod_emp
  End With
  With tabl("emp_act")
    For Each i In labactividad
      If i <> "" Then
        .AddNew Array("cod_emp", "cod_act"), Array(co, i.Tag)
        .Update
      End If
    Next
  End With
  With tabl("emp_cue")
    For Each n In trdisponibles.Nodes
      If n.Checked And n.Children = 0 Then
        .AddNew Array("cod_emp", "cod_cue"), Array(co, Mid(n.key, 2))
        .Update
      End If
    Next
  End With
  If tablaexiste("ingresos" & co & "_copia") And tablaexiste("egresos" & co & "_copia") Then
    If MsgBox("Ya existen tablas de compras y ventas con esta empresa." & vbNewLine & _
              "¿Desea usar las que existen?", vbYesNo) = vbYes Then
      C.Execute "alter table ingresos" & co & "_copia rename to ingresos" & co
      C.Execute "alter table egresos" & co & "_copia rename to egresos" & co
      C.Execute "alter table dingresos" & co & "_copia rename to dingresos" & co
      C.Execute "alter table degresos" & co & "_copia rename to degresos" & co
    Else
      C.Execute "drop table ingresos" & co & "_copia"
      C.Execute "drop table egresos" & co & "_copia"
      C.Execute "drop table dingresos" & co & "_copia"
      C.Execute "drop table degresos" & co & "_copia"
      crearingresos co: crearegresos co
    End If
  End If
  StatusBar1.SimpleText = "Empresa agregada"
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Command6_Click()
  labactividad(act) = ""
  labactividad(act).Tag = ""
End Sub

Private Sub Form_Load()
  initlst lstActividades, Array("Codigo", "Actividad", "Observaciones"), Array(0.1, 0.4, 0.4)
  llenarlst lstActividades, "select * from actividades", Array("cod_act", "nom_act", "obs_act"), "cod_act"
  llenarNivel trdisponibles, "select * from cuentas", "nom_cue", "cod_cue", "cod_pad"
  labactividad_Click 0
  SSTab1.Tab = 0
End Sub

Private Sub labactividad_Click(Index As Integer)
  labactividad(act).BackColor = vbButtonFace
  labactividad(act).ForeColor = vbBlack
  labactividad(Index).BackColor = vbBlack
  labactividad(Index).ForeColor = vbButtonFace
  Command6.top = labactividad(Index).top
  act = Index
End Sub

Private Sub lstactividades_DblClick()
  labactividad(act) = lstActividades.SelectedItem.SubItems(1)
  labactividad(act).Tag = lstActividades.SelectedItem
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  llenarlst lstActividades, "select * from actividades where nom_act like '%" & txtbuscar & "%'", Array("cod_act", "nom_act", "obs_act"), "cod_act"
End Sub

Private Sub trdisponibles_NodeCheck(ByVal Node As Node)
  tildarAbajo Node
  tildarArriba Node
End Sub

Private Sub txtbuscar_Change()
  Timer1.Enabled = True
  Timer1.Interval = 500
End Sub
