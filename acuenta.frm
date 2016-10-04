VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form abmcuenta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtnombre 
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
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   8775
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
      Left            =   4522
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "Agregar"
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
      Left            =   3247
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.TreeView trcuentas 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10398
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   90
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
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
End
Attribute VB_Name = "abmcuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdagregar_Click()
  On Error GoTo E
  assert txtnombre <> "" And trcuentas.SelectedItem <> "", NOCAMP, "Campos obligatorios: nombre y cuenta"
  assert query("cuentas", , "nom_cue='" & txtnombre & "' and cod_pad=" & Mid(trcuentas.SelectedItem.key, 2)).RecordCount = 0, CAMPEX, "La cuenta ya existe"
  With query("cuentas", , "cod_cue=" & Mid(trcuentas.SelectedItem.key, 2))
    !n_hijos = !n_hijos + 1
    .Update
  End With
  With tabl("cuentas")
    .AddNew Array("nom_cue", "cod_pad"), Array(txtnombre, Mid(trcuentas.SelectedItem.key, 2))
    .Update
    trcuentas.Nodes.Add trcuentas.SelectedItem.key, tvwChild, "k" & !cod_cue, !nom_cue
  End With
  StatusBar1.SimpleText = "Cuenta agregada"
  txtnombre = ""
  txtnombre.SetFocus
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  llenarnivel trcuentas, "select * from cuentas", "nom_cue", "cod_cue", "cod_pad"
End Sub

Private Sub cmdeliminar_Click()
  On Error GoTo E
  assert Not trcuentas.SelectedItem Is Nothing, NOCAMP, "Falta seleccionar cuenta"
  If MsgBox("¿Confirma eliminar la cuenta " & trcuentas.SelectedItem & " y todas sus subcuentas?", vbYesNo, "") = vbYes Then
    With query("cuentas", , "cod_cue=" & Mid(trcuentas.SelectedItem.key, 2))
      C.Execute "update cuentas set n_hijos=n_hijos-1 where cod_cue=" & !cod_pad
      .Delete: .Update
    End With
    trcuentas.Nodes.Remove trcuentas.SelectedItem.Index
    StatusBar1.SimpleText = "Cuenta eliminada"
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub trcuentas_AfterLabelEdit(Cancel As Integer, NewString As String)
  With query("cuentas", , "cod_cue=" & Mid(trcuentas.SelectedItem.key, 2))
    !nom_cue = NewString
    .Update
  End With
  StatusBar1.SimpleText = "Cambios guardados"
End Sub

Private Sub trcuentas_DblClick()
  trcuentas.StartLabelEdit
End Sub

Private Sub trcuentas_GotFocus()
  StatusBar1.SimpleText = "Doble clic para editar"
End Sub

Private Sub trcuentas_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

