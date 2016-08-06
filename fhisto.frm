VERSION 5.00
Begin VB.Form fhisto 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7965
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdvolver 
      Caption         =   "Volver"
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
      Left            =   4380
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox pch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
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
      Height          =   6975
      Left            =   120
      ScaleHeight     =   463
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   647
      TabIndex        =   0
      Top             =   480
      Width           =   9735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Totales"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "fhisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
  (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type
Private imeses() As String, emeses() As String, isaldos() As String, esaldos() As String
Private s0 As Double, s1 As Double, m0 As Double, m1 As Double, off As Long
Public im As String, em As String, si As String, se As String

Private Sub cmdvolver_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  centrar Me
  'meses de ingresos y egresos, deberian ser los mismos
  imeses() = Split(im, ",")
  emeses() = Split(em, ",")
  'importes de ingresos y egresos
  isaldos() = Split(si, ",")
  esaldos() = Split(se, ",")
  s0 = isaldos(0): s1 = isaldos(0)
  m0 = ames(imeses(0)): m1 = m0
  For i = 0 To min(UBound(isaldos), UBound(esaldos))
    s0 = min(min(s0, isaldos(i)), esaldos(i))
    s1 = max(max(s1, isaldos(i)), esaldos(i))
    m0 = min(min(m0, ames(imeses(i))), ames(emeses(i)))
    m1 = max(max(m1, ames(imeses(i))), ames(emeses(i)))
  Next
  s0 = Format(min(s0, 0), "0.00")
  s1 = Format(s1, "0.00")
  off = 96
End Sub

Private Sub pch_Paint()
  pch.DrawStyle = vbDot
  For i = off To pch.ScaleHeight - off Step (pch.ScaleHeight - 2 * off) / 20
    pch.Line (off, i)-(pch.ScaleWidth - off, i), &HCCCCCC
  Next
  For i = pch.ScaleWidth - off To off Step -(pch.ScaleWidth - 2 * off) / 20
    pch.Line (i, off)-(i, pch.ScaleHeight - off), &HDDDDDD
  Next
  pch.DrawStyle = vbSolid
  sch = pch.ScaleHeight - off
  scw = pch.ScaleWidth - off
  pch.Line (off, off - 10)-(off, escalary(0))
  pch.Line (off, escalary(0))-(scw + 10, escalary(0))
  pch.DrawWidth = 2
  For i = 0 To UBound(isaldos)
    If i < UBound(isaldos) Then
      pch.Line (escalarx(ames(imeses(i))), escalary(isaldos(i)))- _
               (escalarx(ames(imeses(i + 1))), escalary(isaldos(i + 1))), vbGreen
    End If
    pch.Circle (escalarx(ames(imeses(i))), escalary(isaldos(i))), 3, vbGreen
    escribir escalarx(ames(imeses(i))) - pch.TextWidth(imeses(i)) / 2 - 2, sch + 16, imeses(i), pch
  Next
  For i = 0 To UBound(esaldos)
    If i < UBound(esaldos) Then
      pch.Line (escalarx(ames(emeses(i))), escalary(esaldos(i)))- _
               (escalarx(ames(emeses(i + 1))), escalary(esaldos(i + 1))), vbRed
    End If
    pch.Circle (escalarx(ames(emeses(i))), escalary(esaldos(i))), 3, vbRed
    escribir escalarx(ames(emeses(i))) - pch.TextWidth(emeses(i)) / 2 - 2, sch + 16, emeses(i), pch
  Next
  escribir off - pch.TextWidth("Importe") / 2, off - 20 - pch.TextHeight("Importe"), "Importe", pch
  escribir scw + 24, sch - pch.TextHeight("Periodo") / 2, "Periodo", pch
  If s1 > 0 Then escribir off - 36 - pch.TextWidth(s1) / 2, off - pch.TextHeight(s1) / 2, s1, pch
  escribir off - 36 - pch.TextWidth(s0) / 2, sch - pch.TextHeight(s0) / 2, s0, pch
End Sub

Private Sub escribir(left As Long, top As Long, ByVal str As String, p As Object)
  Dim r As RECT
  r.left = left
  r.right = left + p.TextWidth(str) + 4
  r.top = top
  r.bottom = top + p.TextHeight(str) + 4
  DrawText p.hdc, str, Len(str), r, &H0
End Sub

Private Function escalar(ByVal x As Double, ByVal a_de As Double, ByVal b_de As Double, ByVal a_a As Double, ByVal b_a As Double) As Double
  escalar = a_a + (x - a_de) / IIf(b_de = a_de, 1, b_de - a_de) * (b_a - a_a)
End Function

Private Function escalarx(ByVal x As Double) As Double
  escalarx = escalar(x, m0, m1, off, pch.ScaleWidth - off)
End Function

Private Function escalary(ByVal x As Double) As Double
  escalary = escalar(x, s0, s1, pch.ScaleHeight - off, off)
End Function

