VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form mproveedor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtcodigo 
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
      Left            =   3960
      TabIndex        =   0
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
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
      Left            =   4440
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
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
      Left            =   3960
      TabIndex        =   2
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
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
      Left            =   3120
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtcuit 
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3240
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
      Caption         =   "Cód.proveedor"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
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
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
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
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "mproveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private adoprov As ADODB.Recordset

Private Sub Command4_Click()
  On Error GoTo E
  assert txtcodigo <> "" And Not adoprov Is Nothing, NOCAMP, "Falta ingresar proveedor"
  If MsgBox("¿Realmente desea eliminar el proveedor " & txtnombre & "?", vbYesNo, "") = vbYes Then
    adoprov.Delete
    adoprov.Update
    StatusBar1.SimpleText = "Proveedor eliminado"
    txtcodigo = "": txtcuit = "": txtnombre = "": txtcodigo.SetFocus
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Command5_Click()
  On Error GoTo E
  assert txtnombre <> "" And Not adoprov Is Nothing, NOCAMP, "Ingresar proveedor"
  assert Not adoprov Is Nothing, NOCAMP, "Ingresar proveedor"
  adoprov!nom_prov = txtnombre
  adoprov!cuit_prov = IIf(txtcuit.ClipText = "", Null, txtcuit)
  adoprov.Update
  StatusBar1.SimpleText = "Cambios guardados"
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub txtcodigo_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de proveedor. F3: buscar"
End Sub

Private Sub txtcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    If teclaprov(txtcodigo, txtnombre) Then
      Set adoprov = busc("select * from proveedores where cod_prov=" & txtcodigo)
      txtcuit = coalesce(adoprov!cuit_prov, "")
    Else
      txtcuit = ""
    End If
  End If
End Sub

Private Sub txtcodigo_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtcodigo_Validate(Cancel As Boolean)
  If txtcodigo <> "" Then Cancel = validarprov(txtcodigo, txtnombre)
  If Not Cancel And txtcodigo <> "" Then
    Set adocli = busc("select * from proveedores where cod_prov=" & txtcodigo)
    txtcuit = coalesce(adocli!cuit_prov, "")
  Else
    txtcuit = ""
  End If
End Sub
