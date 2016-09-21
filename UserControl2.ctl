VERSION 5.00
Begin VB.UserControl UserControl2 
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event buscar()
Public Event keydown(keycode As Integer, shift As Integer)
Public Event change()

Private Sub Text1_Change()
  Timer1.Interval = 500
  Timer1.enabled = True
  RaiseEvent change
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.text)
End Sub

Private Sub Text1_KeyDown(keycode As Integer, shift As Integer)
  RaiseEvent keydown(keycode, shift)
End Sub

Private Sub Timer1_Timer()
  Timer1.enabled = False
  RaiseEvent buscar
End Sub

Private Sub UserControl_Resize()
  Text1.Height = UserControl.ScaleHeight
  Text1.Width = UserControl.ScaleWidth
End Sub

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
  Text1.text = PropBag.ReadProperty("text", "")
  Text1.enabled = PropBag.ReadProperty("enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "text", Text1.text, ""
  PropBag.WriteProperty "enabled", Text1.enabled, True
End Sub

