VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form buscar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   -45
   ClientTop       =   -45
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form17"
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
      Left            =   8505
      Picture         =   "buscar.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   2
      Top             =   6360
      Width           =   375
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
      Left            =   105
      TabIndex        =   0
      Top             =   6360
      Width           =   8415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8520
      Top             =   0
   End
   Begin MSComctlLib.ListView lst 
      Height          =   6255
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11033
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
End
Attribute VB_Name = "buscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tabla As String, columna As String, clave As String, busq As String
Public Cancel As Boolean
Public val As Variant, key As Variant
Private sc() As String, sqlc As String

'tabla es donde se buscan los registros
'columna contiene los campos a mostrar en formato sql y separados con el caracter '|'
'clave es el campo que se va a usar para identificar los elementos de la lista
'busq se el campo que se usara para buscar con txtbuscar
'val es el valor de la primera columna del registro seleccionado
'key es la clave del registro seleccionado

Private Sub Form_Load()
  centrar Me
  sc = Split(columna, "|")
  sqlc = Replace(columna, "|", ",")
  For i = 0 To UBound(sc): anc = anc + "," + CStr(1 / (UBound(sc) + 1)): Next
  initlst lst, sc, Split(Mid(anc, 2), ",")
  llenarlst lst, "select " & clave & "," & sqlc & " from " & tabla, sc, clave
End Sub

Private Sub lst_DblClick()
  val = lst.SelectedItem
  key = Mid(lst.SelectedItem.key, 2)
  Cancel = False
  Unload Me
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    val = lst.SelectedItem
    key = Mid(lst.SelectedItem.key, 2)
    Cancel = False
    Unload Me
  End If
End Sub

Private Sub Timer1_Timer()
  Timer1.enabled = False
  llenarlst lst, "select " & clave & "," & sqlc & " from " & tabla & " where " & busq & " like '%" & txtbuscar & "%'", sc, clave
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
