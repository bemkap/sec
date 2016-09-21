VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form livacompras 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9465
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   14880
   ControlBox      =   0   'False
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Columnas"
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
      Left            =   6833
      TabIndex        =   1
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
      Left            =   8153
      TabIndex        =   2
      Top             =   9000
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
      Left            =   5513
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstlistado 
      Height          =   6975
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   12303
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
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   7440
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
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
   Begin MSComctlLib.ListView lstlistado1 
      Height          =   255
      Left            =   13080
      TabIndex        =   6
      Top             =   8880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      View            =   3
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
         Name            =   "Fixedsys"
         Size            =   9
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
      Caption         =   "Listado IVA compras"
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
      TabIndex        =   3
      Top             =   120
      Width           =   14655
   End
End
Attribute VB_Name = "livacompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mostrar() As Variant, ws() As Variant, parcial(9) As Double, parcial1(4) As Double

Private Sub cmdvolver_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
  Dim i As Integer, n As Integer, cw As Double
  With columnas.lstcolumnas
    For i = 1 To lstlistado.ColumnHeaders.Count: .ListItems.Add , , lstlistado.ColumnHeaders(i): Next
    For i = 1 To .ListItems.Count: .ListItems(i).Checked = CBool(mostrar(i - 1)): Next
  End With
  columnas.Show vbModal
  For i = 1 To columnas.v.Count: mostrar(i - 1) = CBool(columnas.v(i)): Next
  For i = 0 To UBound(mostrar): n = n - CInt(mostrar(i)): Next
  cw = lstlistado.Width / n
  For i = 1 To lstlistado.ColumnHeaders.Count
    lstlistado.ColumnHeaders(i).Width = IIf(mostrar(i - 1), cw, 0)
  Next
End Sub

Private Sub Form_Load()
  Dim i As Integer, col1(), col2()
  centrar Me
  ws = Array(10, 5, 14, 15, 13, 8, 12, 8, 9, 8)
  
  For i = 0 To Printer.FontCount - 1
    If Printer.Fonts(i) Like "Courier*" Then
      Printer.Font = Printer.Fonts(i): Exit For
    End If
  Next
  Printer.FontSize = 9
  Printer.ScaleMode = vbCharacters
  
  initlst lstlistado, Array("FECHA", "TIPO", "Nº", "PROVEEDOR", "CUIT", "TOTAL", "EXE-NO GR-IN", "RET./PER", "GRAVADO", "IVA"), _
    Array(0.09, 0.06, 0.11, 0.15, 0.12, 0.09, 0.13, 0.08, 0.08, 0.08)
  initlst lsttotales, Array("C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8"), _
    Array(0.33, 0.2, 0.09, 0.13, 0.08, 0.08, 0.08, 0.1)
  llenarlst lstlistado, _
    "select cod_egr,fecha,nom_comp,numero,nom_prov,cuit_prov1,format(subtotal,'0.00') as fsubtotal," & _
    "format(exento+no_gravado+interno,'0.00') as engi,format(perc_iva+perc_ib,'0.00') as perc," & _
    "format(perc_ib,'0.00') as fperc_ib,format(gravado,'0.00') as fgravado," & _
    "format(iva21+iva105+iva27,'0.00') as giva,cuenta from vegresos order by fecha asc", _
    Array("fecha", "nom_comp", "numero", "nom_prov", "cuit_prov1", "fsubtotal", "engi", "perc", "fgravado", "giva"), "cod_egr"
  initlst lstlistado1, Array("FECHA", "EXENTO", "NO GRAVADO", "INTERNO", "PERC.ING.BTOS", "GRAVADO", "GIVA"), _
    Array(0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1)
  llenarlst lstlistado1, "select fecha,exento,no_gravado,interno,perc_ib,gravado," & _
  "format(iva21+iva105+iva27,'0.00') as giva,iva21,iva105,iva27 from vegresos order by fecha asc", _
    Array("fecha", "exento", "no_gravado", "interno", "perc_ib", "gravado", "giva", "iva21", "iva105", "iva27")
  
  Dim ve As ADODB.Recordset, tasas(2) As ADODB.Recordset
  Set ve = busc("select sum(subtotal),sum(exento+no_gravado+interno),sum(perc_iva+perc_ib),sum(gravado),format(sum(iva21+iva105+iva27),'0.00'),sum(exento),sum(no_gravado),sum(interno),sum(perc_ib) from vegresos")
  Set tasas(0) = busc("select 0.21,sum(format(iva21,'0.00')) from vegresos")
  Set tasas(1) = busc("select 0.27,sum(format(iva27,'0.00')) from vegresos")
  Set tasas(2) = busc("select 0.105,sum(format(iva105,'0.00')) from vegresos")

  With lsttotales.ListItems.Add
    .ListSubItems.Add
    For i = 0 To 4: .ListSubItems.Add , , Format(coalesce(ve.Fields(i), 0), "0.00"): Next
  End With

  col1 = Array("TOTAL EXENTO", "TOTAL NO GRAVADO", "TOTAL INTERNOS", "TOTAL PERC.ING.BTOS")
  col2 = Array("AL 21%", "AL 27%", "AL 10.5%", "")
  For i = 0 To UBound(col1)
    With lsttotales.ListItems.Add
      .ListSubItems.Add , , col1(i)
      .ListSubItems.Add
      .ListSubItems.Add , , Format(coalesce(ve.Fields(i + 5), 0), "0.00")
      .ListSubItems.Add , , col2(i)
      If i < 3 Then .ListSubItems.Add , , Format(coalesce(tasas(i).Fields(1) / tasas(i).Fields(0), 0), "0.00")
      If i < 3 Then .ListSubItems.Add , , Format(coalesce(tasas(i).Fields(1), 0), "0.00")
    End With
  Next
  
  mostrar = Array(True, True, True, True, True, True, True, True, True, True)
  For i = 1 To lstlistado.ColumnHeaders.Count
    If Not mostrar(i - 1) Then lstlistado.ColumnHeaders(i).Width = 0
  Next
End Sub

Private Sub Command2_Click()
  Dim i As Integer, j As Integer, k As Integer, linea As Integer, t As String
  Dim li As ListItem, li1 As ListItem, lij As String
  On Error GoTo E
  selimpr.Show vbModal
  If Not selimpr.cancel Then
    For i = 0 To UBound(parcial): parcial(i) = 0: Next
    For i = 0 To UBound(parcial1): parcial1(i) = 0: Next
    k = 0: titulo k: linea = 11
    For i = 1 To lstlistado.ListItems.Count
      t = ""
      Set li = lstlistado.ListItems(i)
      Set li1 = lstlistado1.ListItems(i)
      t = t & right2(IIf(mostrar(0), li, " "), ws(0)) & " "
      For j = 1 To li.ListSubItems.Count
        lij = li.ListSubItems(j)
        If j >= 5 Then 'numero
          t = t & right2(IIf(mostrar(j), Format(lij, "0.00"), " "), ws(j)) & " "
        Else 'letras
          t = t & left2(IIf(mostrar(j), lij, " "), ws(j)) & " "
        End If
      Next
      For j = 1 To 4: parcial(j - 1) = parcial(j - 1) + li1.ListSubItems(j): Next
      For j = 5 To 9: parcial1(j - 5) = parcial1(j - 5) + li.ListSubItems(j): Next
      parcial(5) = parcial(5) + li1.ListSubItems(7)
      parcial(7) = parcial(7) + li1.ListSubItems(8)
      parcial(9) = parcial(9) + li1.ListSubItems(9)
      yx linea, 1, t
      linea = linea + 1
      If linea > Printer.ScaleHeight - 8 Then
        parciales Printer.ScaleHeight - 7
        Printer.NewPage
        linea = 11
        k = k + 1: titulo k
      End If
    Next
    parciales Printer.ScaleHeight - 7
    Printer.EndDoc
  End If
  Exit Sub
E: MsgBox "Error en la impresión: " & Err.Description, vbCritical, ""
End Sub

Public Sub titulo(ByVal p As Integer)
  Dim t As String, co As ColumnHeader, i As Integer
  yx 1, 1, "HOJA " & (p + 1)
  centro "SUBDIARIO DE IVA COMPRAS DE " & UCase(givacompras.labnom)
  If givacompras.txtfecha(0) <> "  /  /    " Then t = t & " DESDE EL " & givacompras.txtfecha(0)
  If givacompras.txtfecha(1) <> "  /  /    " Then t = t & " HASTA EL " & givacompras.txtfecha(1)
  derecha t
  parciales 2
  For i = 1 To lstlistado.ColumnHeaders.Count
    Set co = lstlistado.ColumnHeaders(i)
    If co.Width = 0 Then co = ""
    t = t & IIf(i >= 6, right2(co, ws(i - 1)), left2(co, ws(i - 1))) & " "
  Next
  yx 9, 1, t
  Printer.Line (1, 10)-(Printer.ScaleWidth - 1, 10)
End Sub

Public Sub parciales(ByVal l As Integer)
  Dim iva(), tit(), i As Integer, t As String
  Printer.Line (1, l)-(Printer.ScaleWidth - 1, l)
  iva = Array(0.21, 0.27, 0.105)
  tit = Array("PARCIAL EXENTO", "PARCIAL NO GRAVADO", "PARCIAL INTERNOS", "PARCIAL PERC.ING.BTOS", "AL 21%", "AL 27%", "AL 10.5%")
  t = String(ws(4) + ws(3) + ws(2) + ws(1) + ws(0), " ") & "    "
  For i = 0 To 4: t = t & right2(Format(parcial1(i), "0.00"), ws(i + 5)) & " ": Next
  yx l + 1, 2, t
  For i = 0 To 3
    t = String(ws(3) + ws(2) + ws(1) + ws(0), " ") & "    "
    t = t & left2(tit(i), ws(5) + ws(4) + 2) & "  " & right2(Format(parcial(i), "0.00"), ws(6) - 2) & " "
    If i < 3 Then
      t = t & left2(tit(i + 4), ws(8)) & right2(Format(parcial(i * 2 + 5) / iva(i), "0.00"), ws(8)) & " "
      t = t & right2(Format(parcial(i * 2 + 5), "0.00"), ws(9))
    End If
    yx l + 2 + i, 1, t
  Next
  Printer.Line (1, l + 6)-(Printer.ScaleWidth - 1, l + 6)
End Sub
