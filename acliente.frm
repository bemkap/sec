VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form abmcliente 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   WindowState     =   2  'Maximized
   Begin Project1.UserControl1 txtcodigo 
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   2925
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      info            =   "Ingresar código de cliente. F3: buscar"
      tabla           =   "clientes"
      campo           =   "nom_cli"
      clave           =   "cod_cli"
      busq            =   "nom_cli"
      regvalid        =   "regvalid"
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
      TabIndex        =   6
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
      Top             =   3645
      Width           =   2895
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
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtcuit 
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3285
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
   Begin VB.Label labcodigo 
      Alignment       =   1  'Right Justify
      Caption         =   "Cód.cliente"
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
      TabIndex        =   7
      Top             =   3045
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Top             =   3405
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   3765
      Width           =   1575
   End
End
Attribute VB_Name = "abmcliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private adocli As ADODB.Recordset
Public alta As Boolean, tmp As Boolean

Private Sub cmdeliminar_Click()
  On Error GoTo E
  assert txtcodigo <> "" And Not adocli Is Nothing, NOCAMP, "Ingresar cliente"
  If MsgBox("¿Realmente desea eliminar el cliente " & txtnombre & "?", vbYesNo, "") = vbYes Then
    adocli!regvalid = False
    adocli!nom_cli = adocli!nom_cli & "(ELIMINADO)"
    adocli.Update
    Set adocli = Nothing
    StatusBar1.SimpleText = "Cliente eliminado"
    txtcodigo = "": txtcuit = "": txtnombre = ""
    txtcodigo.SetFocus
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub cmdguardar_Click()
  On Error GoTo E
  assert txtnombre <> "", NOCAMP, "Campos obligatorios: razón social"
  If alta Then
    Set adocli = tabl("clientes")
    adocli.AddNew
    StatusBar1.SimpleText = "Cliente agregado"
  Else
    assert Not adocli Is Nothing, NOCAMP, "Ingresar cliente"
    StatusBar1.SimpleText = "Cambios guardados"
  End If
  adocli!nom_cli = txtnombre
  adocli!cuit_cli = IIf(txtcuit.ClipText = "", Null, txtcuit)
  adocli.Update
  If alta Then txtcuit.SetFocus Else txtcodigo.SetFocus
  txtcodigo = "": txtcuit = "": txtnombre = ""
  If alta And tmp Then
    tmp = False
    Unload Me
  End If
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Form_Load()
  txtcodigo.Visible = Not alta
  cmdeliminar.Visible = Not alta
  labcodigo.Visible = Not alta
  txtnombre.enabled = alta
  txtcuit.enabled = alta
End Sub

Private Sub txtcodigo_finbusqueda(llave As String, valor As String)
  Set adocli = query("clientes", , "cod_cli=" & llave)
  txtcodigo = llave
  txtnombre = valor
  txtcuit = coalesce(adocli!cuit_cli, "")
  txtnombre.enabled = True
  txtcuit.enabled = True
  txtcuit.SetFocus
End Sub

Private Sub txtcodigo_vacio()
  Set adocli = Nothing
  txtnombre = "": txtnombre.enabled = False
  txtcuit = "": txtcuit.enabled = False
End Sub
