VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form pliva 
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
Attribute VB_Name = "pliva"
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
  Dim hh(), vv(), i As Integer
  pcp.FontBold = True
  pcp.FontSize = 12
  escribir 8, 8, "I.V.A. - AÑO ", pcp
  pcp.FontSize = 10
  pcp.FontBold = False
  escribir 8, 46, "Domicilio fiscal", pcp
  escribir 280, 30, "C.U.I.T.", pcp
  escribir 500, 30, "Act.primaria", pcp
  escribir 500, 46, "Act.secundaria", pcp
  escribir 500, 62, "Act.terciaria", pcp
  pcp.FontSize = 6
  hh = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre") ', "Total")
  vv = Array("D.fiscal(21%)(R.I.)", "D.fiscal(10.5%)(R.I.)", "D.fiscal(21%)(C.F.)", "D.fiscal(10.5%)(C.F.)", "D.fiscal monotributo", "Venta bienes de uso", "Anul.c.fiscal", _
             "C.fiscal(21%)", "C.fiscal(27%)", "C.fiscal(10.5%)", "Compra bienes de uso", "Saldo a favor(art.20 1ºp.)", "Saldo a favor(art.20 2ºp.)", "Retenciones", _
             "Percepciones", "Pagos a cuenta", "D.fiscal(R.N.I.)", "Subtotal", "Saldo libre disp.")
  pcp.Line (8, 82)-(117 + (UBound(hh) + 1) * 66, 82)
  pcp.Line (117, 104)-(117 + (UBound(hh) + 1) * 66, 104)
  For i = 0 To UBound(vv)
    escribir 10, 128 + i * 22, vv(i), pcp
    pcp.Line (8, 126 + i * 22)-(117 + (UBound(hh) + 1) * 66, 126 + i * 22)
  Next
  pcp.Line (8, pcp.ScaleHeight - 14)-(117 + (UBound(hh) + 1) * 66, pcp.ScaleHeight - 14)
  pcp.Line (8, 82)-(8, pcp.ScaleHeight - 14)
  For i = 0 To UBound(hh)
    escribir 119 + i * 66, 82, hh(i), pcp
    escribir 119 + i * 66, 104, "Debe", pcp
    escribir 119 + i * 66 + 33, 104, "Haber", pcp
    pcp.Line (117 + i * 66, 82)-(117 + i * 66, pcp.ScaleHeight - 14)
    pcp.Line (150 + i * 66, 104)-(150 + i * 66, pcp.ScaleHeight - 14)
  Next
  pcp.Line (117 + (UBound(hh) + 1) * 66, 82)-(117 + (UBound(hh) + 1) * 66, pcp.ScaleHeight - 14)
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
    With busc("select * from empresas where cod_emp=" & txtemp)
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
    With busc("select * from emp_act as ea inner join actividades as a on ea.cod_act=a.cod_act " & _
              "where cod_emp=" & txtemp)
      For i = 0 To .RecordCount - 1
        pcp.FontSize = 10
        escribir 610, 30 + i * 16, .Fields("a.cod_act") & "-" & left2(!nom_act, 55), pcp
        .MoveNext
      Next
    End With
  End If
  pcp.FontSize = 12
  pcp.FontBold = True
  escribir 100, 8, txtano, pcp
  pcp.FontBold = False
  pcp.FontSize = 8
End Sub

Private Sub txtemp_vacio()
  labnom = ""
  pcp.Cls
End Sub
