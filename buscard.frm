VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form buscard 
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
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Project1.UserControl2 txtbuscar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8520
      Picture         =   "buscard.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   6480
      Width           =   375
   End
   Begin MSComctlLib.ListView lst 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5106
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
   Begin MSComctlLib.ListView lstd 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
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
Attribute VB_Name = "buscard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public columna As String, clave As String, tabla As String, busq As String
Public detalle As String, excol As String, extab As String, excla As String
Private dt() As String

Private Sub Form_Load()
  initlst lst, Array(columna), Array(1)
  llenarlst lst, "select * from " & tabla, Array(columna), clave
  initlst lstd, Array("Campo", "Valor"), Array(0.3, 0.69)
  dt = Split(detalle, "|")
End Sub

Private Sub lst_DblClick()
  Dim i As Integer
  lstd.ListItems.Clear
  With busc("select " & Replace(detalle, "|", ",") & " from " & tabla & " where " & clave & "=" & Mid(lst.SelectedItem.key, 2))
    For i = 0 To UBound(dt)
      lstd.ListItems.Add(, , ascampo(dt(i))).ListSubItems.Add , , .Fields(ascampo(dt(i)))
    Next
  End With
  lstd.ListItems.Add
  If excol <> "" And extab <> "" And excla <> "" Then
    llenarlst lstd, "select 'Actividad' as c," & excol & "," & excla & " from " & extab & " where " & excla & "=" & Mid(lst.SelectedItem.key, 2), Array("c", excol), , False
  End If
End Sub

Private Sub txtbuscar_buscar()
  llenarlst lst, "select * from " & tabla & " where " & busq & " like '%" & txtbuscar & "%'", Array(columna), clave
End Sub
