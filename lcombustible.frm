VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form lcombustible 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9465
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
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
      Left            =   5828
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
   End
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
      Left            =   7178
      TabIndex        =   1
      Top             =   9000
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstlistado 
      Height          =   8055
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   14208
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
   Begin MSComctlLib.ListView lsttotales 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   8520
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
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
   Begin VB.Label labtitulo 
      Alignment       =   2  'Center
      Caption         =   "Listado gastos de combustible"
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
      TabIndex        =   2
      Top             =   120
      Width           =   13935
   End
End
Attribute VB_Name = "lcombustible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ws() As Variant, parcial(2) As Double

Private Sub cmdvolver_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim i As Integer
  centrar Me
  ws = Array(12, 15, 23, 16, 17, 17)
  
  For i = 0 To Printer.FontCount - 1
    If Printer.Fonts(i) Like "Courier*" Then
      Printer.Font = Printer.Fonts(i): Exit For
    End If
  Next
  Printer.FontSize = 9
  Printer.ScaleMode = vbCharacters
  
  initlst lstlistado, Array("FECHA", "Nº", "PROVEEDOR", "CUIT", "LITROS", "IMPORTE"), _
    Array(0.15, 0.15, 0.2, 0.15, 0.17, 0.17)
  initlst lsttotales, Array("C1", "C2", "C3"), Array(0.65, 0.17, 0.17)
  llenarlst lstlistado, "select cod_egr,fecha,numero,nom_prov,cuit_prov,litros,litros27 from vegresos order by fecha asc", _
    Array("fecha", "numero", "nom_prov", "cuit_prov", "litros", "litros27"), "cod_egr"
  
  Dim ve As ADODB.Recordset
  Set ve = busc("select iif(isnull(sum(litros)),0,sum(litros)),iif(isnull(sum(litros27)),0,sum(litros27)) from vegresos")
  With lsttotales.ListItems.Add
    For i = 0 To 1: .ListSubItems.Add , , ve.Fields(i): Next
  End With
End Sub

Private Sub Command2_Click()
  Dim i As Integer, j As Integer, k As Integer, t As String
  On Error GoTo E
  selimpr.Show vbModal
  If Not selimpr.cancel Then
    k = 0: titulo k
    For i = 1 To lstlistado.ListItems.Count
      t = left2(lstlistado.ListItems(i), ws(0)) & " "
      For j = 1 To lstlistado.ListItems(i).ListSubItems.Count
        t = t & IIf(j >= 4, right2(Format(lstlistado.ListItems(i).ListSubItems(j), "0.00"), ws(j)), _
                            left2(lstlistado.ListItems(i).ListSubItems(j), ws(j))) & " "
      Next
      For j = 4 To 5: parcial(j - 4) = parcial(j - 4) + lstlistado.ListItems(i).ListSubItems(j): Next
      yx i + 6, 4, t
      If i > Printer.ScaleHeight - 3 Then
        parciales Printer.ScaleHeight - 3
        Printer.NewPage
        k = k + 1
        titulo k
        parciales 4
      End If
    Next
    parciales Printer.ScaleHeight - 3
    Printer.EndDoc
  End If
  Exit Sub
E: MsgBox "Error en la impresión: " & Err.Description, vbCritical, ""
End Sub

Public Sub titulo(ByVal p As Integer)
  Dim i As Integer, t As String, co As ColumnHeader
  yx 1, 4, "HOJA " & (p + 1)
  centro UCase(gcombustible.labnom)
  If gcombustible.txtfecha(0) <> "  /  /    " Then t = t & " DESDE EL " & givacompras.txtfecha(0)
  If gcombustible.txtfecha(1) <> "  /  /    " Then t = t & " HASTA EL " & givacompras.txtfecha(1)
  derecha t
  parciales 2
  For i = 1 To lstlistado.ColumnHeaders.Count
    Set co = lstlistado.ColumnHeaders(i)
    t = t & IIf(i >= 5, right2(co, ws(i - 1)), left2(co, ws(i - 1))) & " "
  Next
  yx 5, 4, t
  Printer.Line (4, 6)-(Printer.ScaleWidth - 4, 6)
End Sub

Public Sub parciales(ByVal l As Integer)
  Dim i As Integer, t As String
  Printer.Line (4, l)-(Printer.ScaleWidth - 4, l)
  t = String(ws(2) + ws(1) + ws(0) - 9, " ") & "   PARCIALES" & String(ws(3), " ") & " "
  For i = 0 To 1: t = t & right2(Format(parcial(i), "0.00"), ws(i + 4)) & " ": Next
  yx l + 1, 4, t
  Printer.Line (4, l + 2)-(Printer.ScaleWidth - 4, l + 2)
End Sub
