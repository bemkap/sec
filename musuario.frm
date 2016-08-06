VERSION 5.00
Begin VB.Form musuario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
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
      Left            =   4545
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
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
      Left            =   5160
      TabIndex        =   5
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CheckBox chkpermisos 
      Appearance      =   0  'Flat
      Caption         =   "Impresión de planilla liquidacion"
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
      Left            =   5160
      TabIndex        =   6
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
      Left            =   5160
      TabIndex        =   7
      Top             =   3960
      Width           =   3255
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
      Left            =   720
      TabIndex        =   1
      Top             =   3240
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
      Left            =   720
      TabIndex        =   4
      Top             =   4320
      Width           =   4215
   End
   Begin VB.CommandButton cmdguardar 
      Appearance      =   0  'Flat
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
      Left            =   3225
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
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
      Left            =   720
      TabIndex        =   3
      Top             =   3960
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
      Left            =   720
      TabIndex        =   2
      Top             =   3600
      Width           =   4215
   End
   Begin VB.TextBox txtusuario 
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
      Height          =   360
      Left            =   3465
      TabIndex        =   0
      Top             =   1260
      Width           =   2895
   End
   Begin VB.Label Label17 
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
      Left            =   2280
      TabIndex        =   11
      Top             =   1365
      Width           =   975
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
      TabIndex        =   10
      Top             =   2520
      Width           =   8775
   End
End
Attribute VB_Name = "musuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adousu As ADODB.Recordset

Private Sub cmdeliminar_Click()
  On Error GoTo E
  assert Not adousu Is Nothing, NOCAMP, "Ingresar usuario"
  assert txtusuario <> "admin", INVOP, "admin no se puede eliminar"
  If MsgBox("¿Realmente desea eliminar el usuario " & txtusuario & "?", vbYesNo, "") = vbYes Then
    adousu.Delete
    adousu.Update
    StatusBar1.SimpleText = "Usuario eliminado"
    txtusuario = "": txtusuario.SetFocus
    For i = 0 To chkpermisos.UBound: chkpermisos(i).Value = vbUnchecked: Next
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub cmdguardar_Click()
  On Error GoTo E
  assert Not adousu Is Nothing, NOCAMP, "Falta ingresar usuario"
  assert txtusuario <> "admin", INVOP, "admin no se puede modificar"
  Dim perm As Byte
  For i = 0 To chkpermisos.UBound: perm = perm + chkpermisos(i).Value * 2 ^ i: Next
  adousu!permisos = perm
  adousu!nombre = txtusuario
  adousu.Update
  StatusBar1.SimpleText = "Cambios guardados"
  txtusuario = "": txtusuario.SetFocus
  For i = 0 To chkpermisos.UBound: chkpermisos(i).Value = vbUnchecked: Next
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub txtusuario_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtusuario_Validate False
End Sub

Private Sub txtusuario_Validate(Cancel As Boolean)
  Dim perm As Byte
  Set adousu = busc("select nombre,permisos from usuarios where nombre='" & txtusuario & "'")
  Cancel = adousu.RecordCount <= 0
  StatusBar1.SimpleText = IIf(Cancel, "Usuario inexistente", "")
  If Cancel Then
    For Each i In chkpermisos
      i.Value = vbUnchecked
    Next
  Else
    For i = 0 To chkpermisos.UBound
      If (adousu!permisos And (2 ^ i)) > 0 Then
        chkpermisos(i).Value = vbChecked
      Else
        chkpermisos(i).Value = vbUnchecked
      End If
    Next
  End If
End Sub
