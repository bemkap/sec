VERSION 5.00
Begin VB.UserControl UserControl3 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ScaleHeight     =   735
   ScaleWidth      =   3015
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
Attribute VB_Name = "UserControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_info As String, p_empresa As String, valido As Boolean
Public Event finbusqueda(llave As String, valor As String)
Public Event change()
Public Event vacio()

Private Sub Text1_Change()
  RaiseEvent change
  If Text1 = "" Then RaiseEvent vacio
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.text)
  If Not StatusBar1 Is Nothing Then StatusBar1.SimpleText = p_info
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo E
  Select Case KeyCode
  Case vbKeyF3:
    formbuscar2 p_empresa, "cuentas", "nom_cue", "cod_cue", "cod_pad"
    If Not buscar2.Cancel Then RaiseEvent finbusqueda(buscar2.key, buscar2.val)
  Case vbKeyReturn:
    assert IsNumeric(Text1), INVDAT, "Tipo de código inválido"
    With busc("select * from cuentas where cod_cue=" & Text1)
      If .RecordCount > 0 Then
        If !n_hijos > 0 Then
          StatusBar1.SimpleText = "La cuenta no es usable"
          RaiseEvent vacio
          valido = False
        ElseIf busc("select * from emp_cue where cod_cue=" & Text1 & " and cod_emp=" & p_empresa).RecordCount = 0 Then
          StatusBar1.SimpleText = "La cuenta no está incluída en el plan de cuentas"
          RaiseEvent vacio
          valido = False
        Else
          RaiseEvent finbusqueda(Text1, !nom_cue)
          valido = True
          StatusBar1.SimpleText = ""
        End If
      Else
        StatusBar1.SimpleText = "Cuenta inexistente"
        RaiseEvent vacio
      End If
    End With
  End Select
  Exit Sub
E:
  If Not StatusBar1 Is Nothing Then StatusBar1.SimpleText = Err.Description
End Sub

Private Sub Text1_LostFocus()
  If Not valido Then Text1 = ""
  If Not StatusBar1 Is Nothing Then StatusBar1.SimpleText = ""
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

Public Property Get empresa() As Variant
  empresa = p_empresa
End Property

Public Property Let empresa(ByVal vNewValue As Variant)
  p_empresa = vNewValue
  PropertyChanged "empresa"
End Property

Public Property Get text() As Variant
Attribute text.VB_UserMemId = 0
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
  p_empresa = PropBag.ReadProperty("empresa", "")
  Text1.text = PropBag.ReadProperty("text", "")
  Text1.enabled = PropBag.ReadProperty("enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "info", p_info, ""
  PropBag.WriteProperty "empresa", p_empresa, ""
  PropBag.WriteProperty "text", Text1.text, ""
  PropBag.WriteProperty "enabled", Text1.enabled, True
End Sub

