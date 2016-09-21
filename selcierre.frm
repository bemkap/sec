VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form selcierre 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Guardar"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "COMPRAS"
      TabPicture(0)   =   "selcierre.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstcomp1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstcomp(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdarriba(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdabajo(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "VENTAS"
      TabPicture(1)   =   "selcierre.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdarriba(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdabajo(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lstcomp(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lstcomp1(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdarriba 
         Height          =   375
         Index           =   1
         Left            =   -70440
         Picture         =   "selcierre.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdabajo 
         Height          =   375
         Index           =   1
         Left            =   -71040
         Picture         =   "selcierre.frx":06F2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdabajo 
         Height          =   375
         Index           =   0
         Left            =   3960
         Picture         =   "selcierre.frx":0DAC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdarriba 
         Height          =   375
         Index           =   0
         Left            =   4560
         Picture         =   "selcierre.frx":1466
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   495
      End
      Begin MSComctlLib.ListView lstcomp 
         Height          =   2655
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
      Begin MSComctlLib.ListView lstcomp 
         Height          =   2655
         Index           =   1
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
      Begin MSComctlLib.ListView lstcomp1 
         Height          =   2655
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   3720
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
      Begin MSComctlLib.ListView lstcomp1 
         Height          =   2655
         Index           =   1
         Left            =   -74880
         TabIndex        =   4
         Top             =   3720
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
End
Attribute VB_Name = "selcierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public periodo As Integer, emp As Integer

Private Sub cmdabajo_Click(Index As Integer)
  pasar lstcomp(Index), lstcomp1(Index)
End Sub

Private Sub cmdarriba_Click(Index As Integer)
  pasar lstcomp1(Index), lstcomp(Index)
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  Dim i As ListItem
  For Each i In lstcomp(0).ListItems
    C.Execute "update egresos" & emp & " set periodo=" & periodo & " where cod_egr=" & Mid(i.key, 2)
  Next
  For Each i In lstcomp1(0).ListItems
    C.Execute "update egresos" & emp & " set periodo=0 where cod_egr=" & Mid(i.key, 2)
  Next
  For Each i In lstcomp(1).ListItems
    C.Execute "update ingresos" & emp & " set periodo=" & periodo & " where cod_ing=" & Mid(i.key, 2)
  Next
  For Each i In lstcomp1(1).ListItems
    C.Execute "update ingresos" & emp & " set periodo=0 where cod_ing=" & Mid(i.key, 2)
  Next
  StatusBar1.SimpleText = "Cambios guardados"
End Sub

Private Sub pasar(desde As ListView, hacia As ListView)
  Dim i As ListItem, j As Integer
  For Each i In desde.ListItems
    If i.Selected Then
      With hacia.ListItems.Add(, i.key, i.text)
        .ListSubItems.Add , , i.ListSubItems(1)
        .ListSubItems.Add , , i.ListSubItems(2)
        .ListSubItems.Add , , i.ListSubItems(3)
      End With
    End If
  Next
  For j = desde.ListItems.Count To 1 Step -1
    If desde.ListItems(j).Selected Then desde.ListItems.Remove j
  Next
End Sub

Private Sub Form_Load()
  centrar Me
End Sub
