VERSION 5.00
Begin VB.Form abmusuario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdeliminar 
      Appearance      =   0  'Flat
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
      TabIndex        =   15
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtrepetir 
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
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtnombre 
      Appearance      =   0  'Flat
      DataField       =   "nombre"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox txtclave 
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
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CheckBox chkpermisos 
      Appearance      =   0  'Flat
      Caption         =   "Modificar usuarios"
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
      Left            =   705
      TabIndex        =   3
      Top             =   3240
      Width           =   4215
   End
   Begin VB.CheckBox chkpermisos 
      Appearance      =   0  'Flat
      Caption         =   "Actualizar actividades"
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
      Left            =   705
      TabIndex        =   4
      Top             =   3600
      Width           =   4215
   End
   Begin VB.CheckBox chkpermisos 
      Appearance      =   0  'Flat
      Caption         =   "Registrar empresas, clientes, proveedores"
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
      Left            =   705
      TabIndex        =   5
      Top             =   3960
      Width           =   4215
   End
   Begin VB.CheckBox chkpermisos 
      Appearance      =   0  'Flat
      Caption         =   "Registrar comprobantes"
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
      Left            =   705
      TabIndex        =   6
      Top             =   4320
      Width           =   4215
   End
   Begin VB.CheckBox chkpermisos 
      Appearance      =   0  'Flat
      Caption         =   "Impresión de listados"
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
      Index           =   4
      Left            =   5025
      TabIndex        =   7
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CheckBox chkpermisos 
      Appearance      =   0  'Flat
      Caption         =   "Impresión de planilla liquidación"
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
      Index           =   5
      Left            =   5025
      TabIndex        =   8
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CheckBox chkpermisos 
      Appearance      =   0  'Flat
      Caption         =   "Modificar cuentas"
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
      Index           =   6
      Left            =   5025
      TabIndex        =   9
      Top             =   3960
      Width           =   3255
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
      Left            =   3247
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre"
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
      Left            =   1215
      TabIndex        =   14
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label labclave 
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
      Left            =   1215
      TabIndex        =   13
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label labrepetir 
      Alignment       =   1  'Right Justify
      Caption         =   "Repetir contraseña"
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
      Left            =   1215
      TabIndex        =   12
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label labpermisos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vista de permisos"
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
      Left            =   105
      TabIndex        =   11
      Top             =   2520
      Width           =   8775
   End
End
Attribute VB_Name = "abmusuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public alta As Boolean
Private Hash As New MD5Hash, bytBlock() As Byte, adousu As ADODB.Recordset

Private Sub cmdeliminar_Click()
  Dim i As Integer
  On Error GoTo E
  assert Not adousu Is Nothing, NOCAMP, "Ingresar usuario"
  assert txtnombre <> "admin", INVOP, "admin no se puede eliminar"
  If MsgBox("¿Realmente desea eliminar el usuario " & txtnombre & "?", vbYesNo, "") = vbYes Then
    adousu.Delete
    adousu.Update
    StatusBar1.SimpleText = "Usuario eliminado"
    txtnombre = "": txtnombre.SetFocus
    For i = 0 To chkpermisos.UBound: chkpermisos(i).Value = vbUnchecked: Next
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub cmdguardar_Click()
  Dim i As Integer
  'los permisos se guardan como una suma de potencias de 2
  'son 7: el numero maximo es 1+2+4+8+16+32+64=127
  On Error GoTo E
  Dim perm As Byte
  If alta Then
    assert txtnombre <> "" And txtclave <> "", NOCAMP, "Campos obligatorios: nombre y clave"
    assert txtclave = txtrepetir, INVDAT, "Las contraseñas no coinciden"
    bytBlock = txtclave
    Set adousu = tabl("usuarios")
    adousu.AddNew
    adousu!clave = Hash.HashBytes(bytBlock)
    StatusBar1.SimpleText = "Usuario agregado"
  Else
    assert Not adousu Is Nothing, NOCAMP, "Falta ingresar usuario"
    assert txtnombre <> "admin", INVOP, "admin no se puede modificar"
    StatusBar1.SimpleText = "Cambios guardados"
  End If
  For i = 0 To chkpermisos.UBound: perm = perm + chkpermisos(i).Value * 2 ^ i: Next
  adousu!permisos = perm
  adousu!nombre = txtnombre
  adousu.Update
  txtclave = "": txtrepetir = "": txtnombre = "": txtnombre.SetFocus
  For i = 0 To chkpermisos.UBound: chkpermisos(i).Value = vbUnchecked: Next
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  cmdeliminar.Visible = Not alta
End Sub

Private Sub txtnombre_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not alta And KeyCode = vbKeyReturn Then txtnombre_Validate False
End Sub

Private Sub txtnombre_Validate(Cancel As Boolean)
  Dim i As CheckBox, j As Integer
  If alta Then
    Cancel = busc("select * from usuarios where nombre='" & txtnombre & "'").RecordCount > 0
    StatusBar1.SimpleText = IIf(Cancel, "El usuario ya existe", "")
  Else
    Dim perm As Byte
    Set adousu = busc("select nombre,permisos from usuarios where nombre='" & txtnombre & "'")
    Cancel = adousu.RecordCount <= 0
    StatusBar1.SimpleText = IIf(Cancel, "Usuario inexistente", "")
    If Cancel Then
      For Each i In chkpermisos: i.Value = vbUnchecked: Next
    Else
      For j = 0 To chkpermisos.UBound: chkpermisos(j).Value = ((adousu!permisos And (2 ^ j)) > 0): Next
    End If
  End If
End Sub
