VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtclave 
      Appearance      =   0  'Flat
      DataField       =   "clave"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Entrar"
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtusuario 
      Appearance      =   0  'Flat
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario"
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
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Contraseña"
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
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Left            =   3480
      TabIndex        =   4
      Top             =   3720
      Width           =   3255
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Hash As New MD5Hash, bytBlock() As Byte

Private Sub Command1_Click()
  txtclave_KeyDown vbKeyReturn, 0
End Sub

Private Sub Command2_Click()
  C.Close
  End
End Sub

Private Sub txtclave_Change()
  Label3 = ""
End Sub

Private Sub txtclave_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim reg As ADODB.Recordset
  On Error GoTo E
  assert txtusuario <> "" And txtclave <> "", NOCAMP, ""
  If KeyCode = vbKeyReturn Then
    bytBlock = txtclave
    Set reg = busc("select clave,permisos from usuarios where nombre='" & txtusuario & "'")
    If reg.Fields("clave") = Hash.HashBytes(bytBlock) Then
      p = reg.Fields("permisos")
      U = txtusuario
      inicio.Show
      Unload Me
    Else: GoTo E
    End If
  End If
  Exit Sub
E: Label3 = "Error en usuario o contraseña"
End Sub

Private Sub Form_Load()
  Dim pa As String
  centrar Me
  pa = App.Path & IIf(right(App.Path, 1) = "\", "", "\")
  If Dir(pa & "csv", vbDirectory) = "" Then MkDir (pa & "csv")
  If Dir(pa & "db1.mdb") = "" Then crearbd
  Set C = New ADODB.Connection
  C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pa & "db1.mdb;"
  crearusuarios
  crearactividades
  crearempresas
  crearempcue
  crearempact
  crearcuentas
  crearclientes
  crearproveedores
  crearcomprobantes
End Sub

Private Sub txtusuario_Change()
  Label3 = ""
End Sub
