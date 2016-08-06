VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ScaleHeight     =   735
   ScaleWidth      =   3015
   Begin VB.TextBox Text1 
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
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_info As String, p_tabla As String, p_campo As String, p_clave As String, p_busq As String
Public Event finbusqueda(llave As String, valor As String)
Public Event change()
Public Event alta()
Public Event validate(Cancel As Boolean)

Private Sub Text1_Change()
  RaiseEvent change
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.text)
  If Not StatusBar1 Is Nothing Then StatusBar1.SimpleText = p_info
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If p_tabla <> "" And p_campo <> "" And p_clave <> "" And p_busq <> "" Then
    Select Case KeyCode
    Case vbKeyF3:
      formbuscar p_tabla, p_campo, p_clave, p_busq
      If Not buscar.Cancel Then
        RaiseEvent finbusqueda(buscar.key, buscar.val)
      End If
    Case vbKeyF4:
      RaiseEvent alta
    End Select
  End If
End Sub

Private Sub Text1_LostFocus()
  If Not StatusBar1 Is Nothing Then StatusBar1.SimpleText = ""
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
  RaiseEvent validate(Cancel)
End Sub

Private Sub UserControl_Resize()
  Text1.Height = UserControl.ScaleHeight
  Text1.Width = UserControl.ScaleWidth
End Sub

Public Property Get info() As Variant
  info = p_info
End Property

Public Property Let info(ByVal vNewValue As Variant)
  p_info = vNewValue
  PropertyChanged "info"
End Property

Public Property Get tabla() As Variant
  tabla = p_tabla
End Property

Public Property Let tabla(ByVal vNewValue As Variant)
  p_tabla = vNewValue
  PropertyChanged "tabla"
End Property

Public Property Get campo() As Variant
  campo = p_campo
End Property

Public Property Let campo(ByVal vNewValue As Variant)
  p_campo = vNewValue
  PropertyChanged "campo"
End Property

Public Property Get clave() As Variant
  clave = p_clave
End Property

Public Property Let clave(ByVal vNewValue As Variant)
  p_clave = vNewValue
  PropertyChanged "clave"
End Property

Public Property Get busq() As Variant
  busq = p_busq
End Property

Public Property Let busq(ByVal vNewValue As Variant)
  p_busq = vNewValue
  PropertyChanged "busq"
End Property

Public Property Get text() As Variant
Attribute text.VB_UserMemId = 0
Attribute text.VB_MemberFlags = "200"
  text = Text1.text
End Property

Public Property Let text(ByVal vNewValue As Variant)
  Text1.text = vNewValue
  PropertyChanged "text"
End Property

Public Property Get enabled() As Boolean
  enabled = Text1.enabled
End Property

Public Property Let enabled(ByVal vNewValue As Boolean)
  Text1.enabled = vNewValue
  PropertyChanged "enabled"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  p_info = PropBag.ReadProperty("info", "")
  p_tabla = PropBag.ReadProperty("tabla", "")
  p_campo = PropBag.ReadProperty("campo", "")
  p_clave = PropBag.ReadProperty("clave", "")
  p_busq = PropBag.ReadProperty("busq", "")
  Text1.text = PropBag.ReadProperty("text", "")
  Text1.enabled = PropBag.ReadProperty("enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "info", p_info, ""
  PropBag.WriteProperty "tabla", p_tabla, ""
  PropBag.WriteProperty "campo", p_campo, ""
  PropBag.WriteProperty "clave", p_clave, ""
  PropBag.WriteProperty "busq", p_busq, ""
  PropBag.WriteProperty "text", Text1.text, ""
  PropBag.WriteProperty "enabled", Text1.enabled, True
End Sub
