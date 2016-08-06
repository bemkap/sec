VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form mempresa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12303
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
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
      TabPicture(0)   =   "mempresa.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtcod"
      Tab(0).Control(1)=   "txtdom"
      Tab(0).Control(2)=   "txttel"
      Tab(0).Control(3)=   "txtsus"
      Tab(0).Control(4)=   "txtcar"
      Tab(0).Control(5)=   "txtresp"
      Tab(0).Control(6)=   "txtnom"
      Tab(0).Control(7)=   "txtloc"
      Tab(0).Control(8)=   "txtcuit"
      Tab(0).Control(9)=   "labVis(8)"
      Tab(0).Control(10)=   "labVis(0)"
      Tab(0).Control(11)=   "labVis(1)"
      Tab(0).Control(12)=   "labVis(2)"
      Tab(0).Control(13)=   "labVis(3)"
      Tab(0).Control(14)=   "labVis(4)"
      Tab(0).Control(15)=   "labVis(5)"
      Tab(0).Control(16)=   "labVis(6)"
      Tab(0).Control(17)=   "labVis(7)"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "PLAN DE CUENTAS"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "trdisponibles"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ACTIVIDADES"
      TabPicture(2)   =   "mempresa.frx":001C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "labactividad(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "labactividad(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "labactividad(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lstActividades"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtbuscar"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Command6"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Timer1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdguardar"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdeliminar"
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
         Picture         =   "mempresa.frx":0038
         ScaleHeight     =   330
         ScaleWidth      =   345
         TabIndex        =   31
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox txtcod 
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
         Left            =   -71040
         TabIndex        =   0
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmdeliminar 
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
         Left            =   4560
         TabIndex        =   11
         Top             =   6480
         Width           =   1215
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
         Left            =   3240
         TabIndex        =   10
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   11760
         Top             =   120
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
         Left            =   -71040
         TabIndex        =   3
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txttel 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
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
         Left            =   -71040
         TabIndex        =   5
         Top             =   3480
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
         Left            =   -71040
         TabIndex        =   6
         Top             =   3840
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
         Left            =   -71040
         TabIndex        =   7
         Top             =   4200
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
         Left            =   -71040
         TabIndex        =   8
         Top             =   4560
         Width           =   2895
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
         Left            =   -71040
         TabIndex        =   2
         Top             =   2400
         Width           =   2895
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
         Left            =   -71040
         TabIndex        =   4
         Top             =   3120
         Width           =   2895
      End
      Begin VB.CommandButton Command6 
         Height          =   255
         Left            =   7785
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5520
         Width           =   255
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
         TabIndex        =   13
         Top             =   5040
         Width           =   8775
      End
      Begin MSComctlLib.TreeView trdisponibles 
         Height          =   6255
         Left            =   -74895
         TabIndex        =   9
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   11033
         _Version        =   393217
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
      Begin MSComctlLib.ListView lstActividades 
         Height          =   4575
         Left            =   120
         TabIndex        =   15
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
      Begin MSMask.MaskEdBox txtcuit 
         Height          =   375
         Left            =   -71040
         TabIndex        =   1
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
      Begin VB.Label labVis 
         Alignment       =   1  'Right Justify
         Caption         =   "Cód.empresa"
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
         Index           =   8
         Left            =   -73440
         TabIndex        =   30
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label labVis 
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
         Index           =   0
         Left            =   -73440
         TabIndex        =   29
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label labVis 
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
         Index           =   1
         Left            =   -73440
         TabIndex        =   28
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label labVis 
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
         Index           =   2
         Left            =   -73440
         TabIndex        =   27
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label labVis 
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
         Index           =   3
         Left            =   -73440
         TabIndex        =   26
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label labVis 
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
         Index           =   4
         Left            =   -73440
         TabIndex        =   25
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label labVis 
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
         Index           =   5
         Left            =   -73440
         TabIndex        =   24
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label labVis 
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
         Index           =   6
         Left            =   -73440
         TabIndex        =   23
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label labVis 
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
         Index           =   7
         Left            =   -73440
         TabIndex        =   22
         Top             =   4680
         Width           =   2295
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
         Left            =   945
         TabIndex        =   21
         Top             =   5520
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
         Left            =   945
         TabIndex        =   20
         Top             =   5760
         Width           =   1335
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
         Left            =   945
         TabIndex        =   19
         Top             =   6000
         Width           =   1335
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
         Left            =   2265
         TabIndex        =   18
         Top             =   5520
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
         Left            =   2265
         TabIndex        =   17
         Top             =   5760
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
         Index           =   2
         Left            =   2265
         TabIndex        =   16
         Top             =   6000
         Width           =   5775
      End
   End
End
Attribute VB_Name = "mempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim act As Integer, adoemp As ADODB.Recordset

Private Sub cmdguardar_Click()
  On Error GoTo E
  assert Not adoemp Is Nothing, NOCAMP, "Falta elegir empresa"
  Dim co As Integer, n As Node
  C.Execute "delete from emp_act where cod_emp=" & txtcod
  C.Execute "delete from emp_cue where cod_emp=" & txtcod
  With adoemp
    .Update Array("cuit_emp", "nom_emp", "dom_emp", "loc_emp", "tel_emp", "sus_emp", "car_emp", "resp_emp"), _
            Array(txtcuit, txtnom, txtdom, txtloc, txttel, txtsus, txtcar, txtresp)
  End With
  With tabl("emp_act")
    For Each i In labactividad
      If i <> "" Then
        .AddNew Array("cod_emp", "cod_act"), Array(txtcod, i.Tag)
        .Update
      End If
    Next
  End With
  With tabl("emp_cue")
    For Each n In trdisponibles.Nodes
      If n.Checked And n.Children = 0 Then
        .AddNew Array("cod_emp", "cod_cue"), Array(adoemp!cod_emp, Mid(n.key, 2))
        .Update
      End If
    Next
  End With
  StatusBar1.SimpleText = "Cambios guardados"
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

Private Sub cmdeliminar_Click()
  If MsgBox("¿Desea guardar una copia de los movimientos de la empresa?", vbYesNo, "") = vbYes Then
    C.Execute "alter table ingresos" & txtcod & " rename to ingresos" & txtcod & "_copia"
    C.Execute "alter table egresos" & txtcod & " rename to egresos" & txtcod & "_copia"
    C.Execute "alter table dingresos" & txtcod & " rename to dingresos" & txtcod & "_copia"
    C.Execute "alter table degresos" & txtcod & " rename to degresos" & txtcod & "_copia"
  Else
    C.Execute "drop table ingresos" & txtcod
    C.Execute "drop table egresos" & txtcod
    C.Execute "drop table dingresos" & txtcod
    C.Execute "drop table degresos" & txtcod
  End If
  C.Execute "delete from empresas where cuit_emp=" & txtcod
  C.Execute "delete from emp_act where cuit_emp=" & txtcod
  C.Execute "delete from emp_sal where cuit_emp=" & txtcod
  C.Execute "delete from emp_cue where cuit_emp=" & txtcod
  StatusBar1.SimpleText = "Empresa eliminada"
End Sub

Private Sub txtcod_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de empresa. F3: buscar"
End Sub

Private Sub txtcod_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    If teclaemp(txtcod, txtnom) Then llenaremp
  End If
End Sub

Private Sub txtcod_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtcod_Validate(Cancel As Boolean)
  If txtcod <> "" Then Cancel = validaremp(txtcod, txtnom)
  If Cancel Then
    txtcuit = "": txtdom = "": txtloc = "": txttel = "": txtsus = "": txtcar = "": txtresp = ""
    For Each n In trdisponibles.Nodes: n.Checked = False: Next
    For i = 0 To 2: labactividad(i) = "": labactividad_Click 0: Next
  Else
    llenaremp
  End If
End Sub

Private Sub llenaremp()
  Set adoemp = busc("select * from empresas where cod_emp=" & txtcod)
  With adoemp
    txtcuit = !cuit_emp
    txtdom = !dom_emp
    txtloc = !loc_emp
    txttel = !tel_emp
    txtsus = !sus_emp
    txtcar = !car_emp
    txtresp = !resp_emp
  End With
  With busc("select * from emp_cue where cod_emp=" & txtcod)
    Do While Not .EOF
      Dim n As Node: Set n = trdisponibles.Nodes("k" & !cod_cue)
      n.Checked = True
      Do While Not n Is Nothing
        tildarArriba n
        Set n = n.Parent
      Loop: .MoveNext
    Loop
  End With
  With busc("select actividades.cod_act, actividades.nom_act from actividades " & _
            "inner join emp_act on emp_act.cod_act=actividades.cod_act " & _
            "where emp_act.cod_emp=" & txtcod)
    i = 0
    Do Until .EOF
      labactividad(i) = !nom_act
      labactividad(i).Tag = !cod_act
      .MoveNext: i = i + 1
    Loop
  End With
End Sub

