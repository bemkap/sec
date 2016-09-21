Attribute VB_Name = "misc"
Option Explicit
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public hijo As Object
Public StatusBar1 As StatusBar

Public Sub centrar(frm As Form)
  frm.top = (Screen.Height - frm.Height) / 2
  frm.left = (Screen.Width - frm.Width) / 2
End Sub

Public Sub abrir(ByVal par As Object, ByVal chld As Object, Optional cerrar As Boolean = True)
  If Not hijo Is Nothing And cerrar Then Unload hijo
  Load chld
  SetParent chld.hWnd, par.hWnd
  chld.Show
  Set hijo = chld
End Sub

Public Sub formbuscard(ByVal par As Object, tabla As String, campo As String, clave As String, busq As String, det As String)
  buscard.tabla = tabla
  buscard.columna = campo
  buscard.clave = clave
  buscard.busq = busq
  buscard.detalle = det
  abrir par, buscard
End Sub

Public Function cjoin(C As Collection, d As String) As String
  Dim i As Integer
  cjoin = ""
  If C.Count = 0 Then Exit Function
  For i = 2 To C.Count
    If C.item(i) <> "" Then cjoin = cjoin & d & C.item(i)
  Next
  cjoin = IIf(C.item(1) <> "", C.item(1) & cjoin, Mid(cjoin, Len(d)))
End Function

Public Function borden(C As String, d0 As String, d1 As String) As String
  borden = IIf(C = "", "", d0 + C + d1)
End Function

Public Function ascampo(ByVal sql As String) As String
  Dim n As Integer
  n = InStr(1, sql, " as ")
  If n > 0 And Len(sql) > 3 Then ascampo = right(sql, Len(sql) - n - 3) Else ascampo = sql
End Function

Public Function min(ByVal x As Double, ByVal y As Double) As Double
  min = IIf(x < y, x, y)
End Function

Public Function max(ByVal x As Double, ByVal y As Double) As Double
  max = IIf(x > y, x, y)
End Function

Public Function ames(ByVal aniomes As String) As Integer
  Dim am()
  am = Split(aniomes, "/")
  ames = am(1) * 12 + am(0)
End Function

Public Function left2(ByVal str As String, ByVal i As Integer) As String
  left2 = left(str & String(i, " "), i)
End Function

Public Function right2(ByVal str As String, ByVal i As Integer) As String
  right2 = right(String(i, " ") & str, i)
End Function

Public Function coalesce(ParamArray args() As Variant) As Variant
  Dim i As Integer
  For i = 0 To UBound(args)
    If Not IsNull(args(i)) Then
      coalesce = args(i)
      Exit Function
    End If
  Next
End Function
