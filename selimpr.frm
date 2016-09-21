VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form selimpr 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2955
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
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
      Left            =   3660
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
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
      Left            =   2340
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstimpresoras 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      View            =   3
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
Attribute VB_Name = "selimpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private colimpr As New Collection
Public Cancel As Boolean

Private Sub Command1_Click()
  Cancel = True
  Unload Me
End Sub

Private Sub Command2_Click()
  Set Printer = colimpr(lstimpresoras.SelectedItem.Index)
  Unload Me
End Sub

Private Sub Form_Load()
  Dim p As Printer
  centrar Me
  Cancel = False
  lstimpresoras.ColumnHeaders.Add , , "Impresoras", lstimpresoras.Width
  For Each p In Printers
    colimpr.Add p
    lstimpresoras.ListItems.Add , p.DeviceName, p.DeviceName
  Next
End Sub

Private Sub lstimpresoras_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Command2_Click
End Sub
