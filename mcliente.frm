VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form mcliente 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form14"
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
   Begin VB.CommandButton Command1 
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
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
   Begin VB.Label labVis 
      Alignment       =   1  'Right Justify
      Caption         =   "Cód.clinte"
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
      TabIndex        =   7
      Top             =   3000
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
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   3720
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
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "mcliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private adocli As ADODB.Recordset

Private Sub Command1_Click()
  On Error GoTo E
  assert txtcodigo <> "" And Not adocli Is Nothing, NOCAMP, "Ingresar cliente"
  If MsgBox("¿Realmente desea eliminar el cliente " & txtnombre & "?", vbYesNo, "") = vbYes Then
    adocli.Delete
    adocli.Update
    StatusBar1.SimpleText = "Cliente eliminado"
    txtcodigo = "": txtcuit = "": txtnombre = ""
    txtcodigo.SetFocus
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Command2_Click()
  On Error GoTo E
  assert txtnombre <> "" And Not adocli Is Nothing, NOCAMP, "Ingresar cliente"
  assert Not adocli Is Nothing, NOCAMP, "Ingresar cliente"
  adocli!nom_cli = txtnombre
  adocli!cuit_cli = IIf(txtcuit.ClipText = "", Null, txtcuit)
  adocli.Update
  StatusBar1.SimpleText = "Cambios guardados"
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub txtcodigo_GotFocus()
  StatusBar1.SimpleText = "Ingresar código de cliente. F3: buscar"
End Sub

Private Sub txtcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    If teclacli(txtcodigo, txtnombre) Then
      Set adocli = busc("select * from clientes where cod_cli=" & txtcodigo)
      txtcuit = coalesce(adocli!cuit_cli, "")
    End If
  End If
End Sub

Private Sub txtcodigo_LostFocus()
  StatusBar1.SimpleText = ""
End Sub

Private Sub txtcodigo_Validate(Cancel As Boolean)
  If txtcodigo <> "" Then Cancel = validarcli(txtcodigo, txtnombre)
  If Not Cancel And txtcodigo <> "" Then
    Set adocli = busc("select * from clientes where cod_cli=" & txtcodigo)
    txtcuit = coalesce(adocli!cuit_cli, "")
  End If
End Sub
