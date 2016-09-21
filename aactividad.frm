VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form aactividad 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recuperar"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Abrir"
      Filter          =   "Text Files (*.txt)|*.txt"
      FilterIndex     =   1
   End
   Begin MSComctlLib.ListView lstactividades 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10398
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
   Begin VB.CommandButton Command4 
      Caption         =   "Cargar"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   6480
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
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   8775
   End
End
Attribute VB_Name = "aactividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  If tablaexiste("actividades_cds") Then
    If tabl("actividades_cds").RecordCount > 0 Then
      llenarlst lstactividades, "select * from actividades_cds", Array("cod_act", "nom_act", "obs_act"), "cod_act"
      StatusBar1.SimpleText = lstactividades.ListItems.Count & " actividades cargadas"
    Else
      StatusBar1.SimpleText = "No hay actividades guardadas"
    End If
  End If
End Sub

Private Sub Command3_Click()
  On Error GoTo E
  Dim i As ListItem
  If MsgBox("Al actualizar las actividades se borrarán las relaciones entre empresas y actividades. Continuar?", vbYesNo) = vbYes Then
    C.Execute "delete from emp_act"
    C.Execute "drop table actividades_cds"
    C.Execute "select * into actividades_cds from actividades"
    C.Execute "delete from actividades"
    With tabl("actividades")
      For Each i In lstactividades.ListItems
        .AddNew Array("cod_act", "nom_act", "obs_act"), Array(i, i.ListSubItems(1), i.ListSubItems(2))
      Next
      .Update
    End With
    StatusBar1.SimpleText = "Actividades actualizadas"
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Command4_Click()
  On Error GoTo E
  Dim line As String, camp() As String, n As Integer
  dialog.ShowOpen
  lstactividades.ListItems.Clear
  labactividad = dialog.FileName
  Open dialog.FileName For Input As #1
  Do Until EOF(1)
    Line Input #1, line
    camp = Split(line, "@")
    With lstactividades.ListItems.Add()
      .text = camp(0)
      .ListSubItems.Add , , Mid(camp(1), 1, 200)
      .ListSubItems.Add , , Mid(camp(2), 1, 200)
    End With
    n = n + 1
  Loop
  Close #1
  StatusBar1.SimpleText = (n & " actividades cargadas")
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  dialog.InitDir = App.Path
  initlst lstactividades, Array("Código", "Actividad", "Observaciones"), Array(0.1, 0.4, 0.4)
End Sub

