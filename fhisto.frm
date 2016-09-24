VERSION 5.00
Begin VB.Form fhisto 
   AutoRedraw      =   -1  'True
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
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
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
Option Explicit
Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
  (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private imeses(11) As Integer, emeses(11) As Integer, isaldos(11) As Double, esaldos(11) As Double
Private s0 As Double, s1 As Double, off As Long, ne As Integer, ni As Integer
Public aa As Double

Private Sub cmdvolver_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim i As Integer
  centrar Me
  s0 = 0: s1 = 0
  With busc("select periodo,sgravado+sno_gravado+siva+sexento+sinterno+sperc_iva+sperc_ib+slitros" & _
    " from vte order by periodo asc")
    ne = .RecordCount
    For i = 0 To .RecordCount - 1
      esaldos(.Fields(0) - aa * 12) = .Fields(1)
      s0 = min(s0, .Fields(1))
      s1 = max(s1, .Fields(1))
      .MoveNext
    Next
  End With
  With busc("select periodo,sgravado+sno_gravado+siva+sexento+sinterno+sret_iva+sret_ib" & _
    " from vti order by periodo asc")
    ni = .RecordCount
    For i = 0 To .RecordCount - 1
      isaldos(.Fields(0) - aa * 12) = .Fields(1)
      s0 = min(s0, .Fields(1))
      s1 = max(s1, .Fields(1))
      .MoveNext
    Next
  End With
  s0 = Format(s0, "0.00")
  s1 = Format(s1, "0.00")
  off = 96
  pch_Paint
End Sub

Private Sub pch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer, cuad As IPictureDisp
  Dim ex As Double, ey As Double, fs As String
  pch.Cls
  pch_Paint
  pch.ForeColor = vbWhite
  Set cuad = LoadPicture("iconos/rect.bmp")
  For i = 0 To UBound(isaldos) - 1
    ex = escalarx(i)
    ey = escalary(isaldos(i))
    fs = "$" & Format(isaldos(i), "0.00")
    If x > ex - 10 And x < ex + 10 And y > ey - 10 And y < ey + 10 Then
      pch.PaintPicture cuad, ex, ey - 20, pch.TextWidth(fs), pch.TextHeight(fs)
      escribir ex, ey - 20, fs
    End If
  Next
  For i = 0 To UBound(esaldos) - 1
    ex = escalarx(i)
    ey = escalary(esaldos(i))
    fs = "$" & Format(esaldos(i), "0.00")
    If x > ex - 10 And x < ex + 10 And y > ey - 10 And y < ey + 10 Then
      pch.PaintPicture cuad, ex, ey - 20, pch.TextWidth(fs), pch.TextHeight(fs)
      escribir ex, ey - 20, fs
    End If
  Next
End Sub

Private Sub pch_Paint()
  Dim i As Double, x As Double, y As Double, xx As Double, yy As Double
  Dim ex As Double, ex1 As Double, ey As Double, ey1 As Double
  Dim sch As Integer, scw As Integer
  pch.Font = "Courier"
  pch.FontSize = 10
  'grilla
  pch.DrawWidth = 1
  pch.ForeColor = vbBlack
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
  'ejes
  pch.Line (off, off - 10)-(off, escalary(0))
  pch.Line (off, escalary(0))-(scw + 10, escalary(0))
  pch.DrawWidth = 3
  'lineas ingreso
  For i = 0 To UBound(isaldos) - 1
    ex = escalarx(i): ex1 = escalarx(i + 1)
    ey = escalary(isaldos(i)): ey1 = escalary(isaldos(i + 1))
    pch.Line (ex, ey)-(ex1, ey1), &H119900
    pch.Circle (ex, ey), 3, &H119900
    pch.Circle (ex, ey), 1, vbWhite
  Next
  'lineas egreso
  For i = 0 To UBound(esaldos) - 1
    ex = escalarx(i): ex1 = escalarx(i + 1)
    ey = escalary(esaldos(i)): ey1 = escalary(esaldos(i + 1))
    pch.Line (ex, ey)-(ex1, ey1), vbRed
    pch.Circle (ex, ey), 3, vbRed
    pch.Circle (ex, ey), 1, vbWhite
  Next
  'periodos
  pch.FontBold = True
  escribir scw + 24, sch - pch.TextHeight("Periodo") / 2, "Periodo"
  For i = 0 To 11: escribir escalarx(i) - pch.TextWidth(i) / 2 - 2, sch + 16, i + 1: Next
  'importes
  escribir off - pch.TextWidth("Importe") / 2, off - 20 - pch.TextHeight("Importe"), "Importe"
  If s1 > 0 Then escribir off - 48 - pch.TextWidth(s1) / 2, off - pch.TextHeight(s1) / 2, s1
  escribir off - 36 - pch.TextWidth(s0) / 2, sch - pch.TextHeight(s0) / 2, s0
  'totales anuales
  pch.FontSize = 12
  pch.ForeColor = &H119900
  escribir_centro 10, "Total ventas  " & Format(busc("select sum(sgravado+sno_gravado+siva+sexento+sinterno+sret_iva+sret_ib) from vti").Fields(0), "0.00")
  pch.ForeColor = vbRed
  escribir_centro 32, "Total compras " & Format(busc("select sum(sgravado+sno_gravado+siva+sexento+sinterno+sperc_iva+sperc_ib+slitros) from vte").Fields(0), "0.00")
  'año
  pch.ForeColor = vbBlack
  escribir_centro pch.ScaleHeight - off / 2, "Año " & aa
End Sub

Private Sub escribir(ByVal left As Long, ByVal top As Long, ByVal str As String)
  Dim r As RECT
  r.left = left
  r.right = left + pch.TextWidth(str) + 4
  r.top = top
  r.bottom = top + pch.TextHeight(str) + 4
  DrawText pch.hdc, str, Len(str), r, &H0
End Sub

Private Sub escribir_centro(ByVal top As String, ByVal str As String)
  escribir (pch.ScaleWidth - pch.TextWidth(str)) / 2, top, str
End Sub

Private Function escalar(ByVal x As Double, ByVal a_de As Double, ByVal b_de As Double, ByVal a_a As Double, ByVal b_a As Double) As Double
  escalar = a_a + (x - a_de) / IIf(b_de = a_de, 1, b_de - a_de) * (b_a - a_a)
End Function

Private Function escalarx(ByVal x As Double) As Double
  escalarx = escalar(x, 0, 11, off, pch.ScaleWidth - off)
End Function

Private Function escalary(ByVal x As Double) As Double
  escalary = escalar(x, s0, s1, pch.ScaleHeight - off, off)
End Function

