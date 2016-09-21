VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form buscar2 
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8520
      Picture         =   "buscar2.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   6480
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   8415
   End
   Begin MSComctlLib.TreeView tr 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11245
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   90
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
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
   Begin MSComctlLib.TreeView tr1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   90
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
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
Attribute VB_Name = "buscar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codemp As String, tabla As String, columna As String, clave As String, padre As String
Public Cancel As Boolean
Public val As Variant, key As Variant

'codemp es el codigo de empresa
'tabla es donde se buscan los registros
'columna es el campo que se mostrar en los nodos
'clave es el campo que identificara al los nodos
'padre es para la contruccion del arbol
'val es el valor del nodo seleccionado
'key es la clave del nodo seleccionado

Private Sub Form_Load()
  centrar Me
  llenarnivel tr, "select * from cuentas where n_hijos>0", columna, clave, padre
  llenarnivel tr, "select emp_cue.cod_cue,emp_cue.cod_emp,cuentas.nom_cue,cuentas.cod_pad " & _
                  "from emp_cue inner join cuentas on emp_cue.cod_cue=cuentas.cod_cue " & _
                  "where emp_cue.cod_emp=" & codemp, _
                  columna, clave, padre, False
  'se tienen 2 arboles para la busqueda
  llenarnivel tr1, "select * from cuentas where n_hijos>0", columna, clave, padre
  llenarnivel tr1, "select emp_cue.cod_cue,emp_cue.cod_emp,cuentas.nom_cue,cuentas.cod_pad " & _
                   "from emp_cue inner join cuentas on emp_cue.cod_cue=cuentas.cod_cue " & _
                   "where emp_cue.cod_emp=" & codemp, _
                   columna, clave, padre, False
End Sub

Private Sub tr_DblClick()
  'no se pueden elegir nodos que tengan subnodos
  If busc("select * from cuentas where cod_cue=" & Mid(tr.SelectedItem.key, 2)).Fields("n_hijos") = 0 Then
    val = tr.SelectedItem
    key = Mid(tr.SelectedItem.key, 2)
    Cancel = False
    Unload Me
  End If
End Sub

Private Sub tr_KeyDown(KeyCode As Integer, Shift As Integer)
  tr_DblClick
End Sub

Private Sub Timer1_Timer()
  Dim n As Node
  If columna <> "" And clave <> "" And tabla <> "" And padre <> "" Then
    Timer1.enabled = False
    tr.Nodes.Clear
    For Each n In tr1.Nodes
      If n.Children > 0 Or InStr(1, n, txtbuscar, 1) > 0 Then
        If n.Parent Is Nothing Then tr.Nodes.Add , , n.key, n Else tr.Nodes.Add n.Parent.key, tvwChild, n.key, n
      End If
    Next
  End If
End Sub

Private Sub txtbuscar_Change()
  Timer1.Interval = 500
  Timer1.enabled = True
End Sub

Private Sub txtbuscar_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Cancel = True
    Unload Me
  End If
End Sub

