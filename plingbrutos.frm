VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form plingbrutos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9825
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   14220
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
      Left            =   7163
      TabIndex        =   3
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
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
      Left            =   5843
      TabIndex        =   2
      Top             =   9360
      Width           =   1215
   End
   Begin VB.PictureBox pcp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8445
      Left            =   120
      ScaleHeight     =   561
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   927
      TabIndex        =   5
      Top             =   840
      Width           =   13935
   End
   Begin MSMask.MaskEdBox txtano 
      Height          =   375
      Left            =   11010
      TabIndex        =   1
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin Project1.UserControl1 txtemp 
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      info            =   "Ingresar código de empresa. F3: buscar"
      tabla           =   "empresas"
      campo           =   "nom_emp"
      clave           =   "cod_emp"
      busq            =   "nom_emp"
      regvalid        =   "regvalid"
   End
   Begin VB.Label labnom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label17 
      Caption         =   "Cód.emp"
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
      Left            =   2610
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Año"
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
      Left            =   10170
      TabIndex        =   4
      Top             =   345
      Width           =   615
   End
End
Attribute VB_Name = "plingbrutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
  (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Private Sub cmdvolver_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  On Error GoTo E
  assert txtano <> "" And txtemp <> "", NOCAMP, "Campos obligatorios: empresa y año"
  selimpr.Show vbModal
  If Not selimpr.Cancel Then
    Printer.Orientation = vbPRORLandscape
    Printer.PaperSize = vbPRPSLegal
    Printer.PaintPicture pcp.Image, 0, 0, Printer.Width, pcp.Height / pcp.Width * Printer.Width
    Printer.EndDoc
  End If
  Exit Sub
E: MsgBox Err.Description: Printer.KillDoc
End Sub

Private Sub Form_Load()
  centrar Me
  pcp_Paint
End Sub

Private Sub escribir(left As Long, top As Long, ByVal str As String, p As PictureBox)
  Dim r As RECT
  r.left = left
  r.right = left + p.TextWidth(str)
  r.top = top
  r.bottom = top + p.TextHeight(str)
  DrawText p.hdc, str, Len(str), r, &H0
End Sub

Private Sub pcp_Paint()
  Dim vv(), vh(), hh0, hh1, i As Integer
  pcp.FontBold = True
  pcp.FontSize = 12
  escribir 8, 8, "IMPUESTO INGRESOS BRUTOS - AÑO ", pcp
  pcp.FontSize = 10
  pcp.FontBold = False
  escribir 8, 46, "Domicilio fiscal", pcp
  escribir 280, 30, "C.U.I.T.", pcp
  escribir 500, 30, "Act.primaria", pcp
  escribir 500, 46, "Act.secundaria", pcp
  escribir 500, 62, "Act.terciaria", pcp
  pcp.FontSize = 6
  vv = Array("Per.", "Ene.", "Feb.", "Mar.", "Abr.", "May.", "Jun.", "Jul.", "Ago.", "Sep.", "Oct.", "Nov.", "Dic.", "Total")
  vh = Array(0, 60, 60, 60, 55, 50, 40, 55, 55, 55, 45, 60, 68, 40, 40, 40, 40)
  hh0 = Array("Monto impon.", "Monto impon.", "Monto impon.", "Impuesto", "Diferencia", "Subtotal", "Saldo a", "Retenciones", "Percep. del", "D.R.I.", "Saldo a", "Saldo a pagar", "Alícuota", "Alícuota", "Subtotal", "Alícuota", "Saldo a pagar")
  hh1 = Array("Alic:.........%", "Alic:.........%", "Alic:.........%", "determinado", "al mínimo", "", "fav.per.ant.", "del periodo", "periodo", "", "fav.prox.per", "", ".........%", ".........%", "", "10%", "")
  pcp.Line (8, 98)-(pcp.ScaleWidth - 8, 98)
  pcp.Line (8, 82)-(pcp.ScaleWidth - 8, 82)
  For i = 0 To UBound(vv)
    escribir 8, 116 + i * 32, vv(i), pcp
  Next
  For i = 0 To UBound(vv) * 2
    pcp.Line (8, 132 + i * 16)-(pcp.ScaleWidth - 8, 132 + i * 16)
  Next
  pcp.Line (7, 82)-(7, pcp.ScaleHeight - 12)
  Dim h As Long: h = 32
  For i = 0 To min(UBound(hh0), UBound(hh1))
    h = h + vh(i)
    escribir h + 2, 98, hh0(i), pcp
    escribir h + 2, 116, hh1(i), pcp
    pcp.Line (h, 82)-(h, pcp.ScaleHeight - 12)
    If i = 12 Then pcp.Line (h - 3, 82)-(h - 3, pcp.ScaleHeight - 12)
  Next
  escribir 817, 84, "Cartel.", pcp
  pcp.Line (pcp.ScaleWidth - 9, 82)-(pcp.ScaleWidth - 9, pcp.ScaleHeight - 12)
End Sub

Private Sub txtano_GotFocus()
  txtano.SelStart = 0
  txtano.SelLength = 4
End Sub

Private Sub txtano_LostFocus()
  pcp.Cls: pcp_Paint: llenardat
End Sub

Private Sub txtemp_finbusqueda(llave As String, valor As String)
  txtemp = llave
  labnom = valor
  pcp.Cls: pcp_Paint: llenardat
  txtano.SetFocus
End Sub

Private Sub llenardat()
  Dim i As Integer
  If txtemp <> "" Then
    With query("empresas", , "cod_emp=" & txtemp)
      pcp.FontBold = True
      pcp.FontSize = 12
      escribir 610, 8, UCase(!resp_emp), pcp
      pcp.FontSize = 10
      pcp.FontBold = False
      escribir 8, 30, StrConv(!car_emp, VbStrConv.vbProperCase), pcp
      escribir 110, 30, StrConv(!sus_emp, VbStrConv.vbProperCase), pcp
      escribir 110, 46, StrConv(!dom_emp, VbStrConv.vbProperCase), pcp
      escribir 350, 30, Format(!cuit_emp, "00-00000000-0"), pcp
    End With
    With query("emp_act as ea inner join actividades as a on ea.cod_act=a.cod_act", , "cod_emp=" & txtemp)
      For i = 0 To .RecordCount - 1
        pcp.FontSize = 10
        escribir 610, 30 + i * 16, .fields("a.cod_act") & "-" & left2(!nom_act, 55), pcp
        pcp.FontSize = 8
        escribir 34 + i * 60, 84, "A" & i + 1 & " : " & .fields("a.cod_act"), pcp
        .MoveNext
      Next
    End With
  End If
  pcp.FontSize = 12
  pcp.FontBold = True
  escribir 310, 8, txtano, pcp
  pcp.FontBold = False
  pcp.FontSize = 8
End Sub

Private Sub txtemp_vacio()
  labnom = ""
  pcp.Cls
End Sub
