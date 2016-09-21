Attribute VB_Name = "arbol"
Option Explicit

Public Sub tildarabajo(ByVal nodo As Node)
  Dim odon As Node, i As Integer
  If nodo.Children > 0 Then
    Set odon = nodo.Child
    For i = 1 To nodo.Children
      odon.Checked = nodo.Checked
      odon.Tag = nodo.Checked
      tildarabajo odon
      Set odon = odon.Next
    Next
  End If
End Sub

Public Sub tildararriba(ByVal nodo As Node)
  Dim odon As Node, i As Integer
  If Not nodo.Parent Is Nothing Then
    Set odon = nodo.FirstSibling
    nodo.Parent.Checked = True
    nodo.Parent.Tag = 0
    For i = 1 To odon.Parent.Children
      nodo.Parent.Checked = nodo.Parent.Checked And odon.Checked
      nodo.Parent.Tag = nodo.Parent.Tag Or odon.Checked
      Set odon = odon.Next
    Next
  End If
End Sub

Public Sub llenarnivel(tr As TreeView, sql As String, col As String, key As String, pad As String, Optional vaciar As Boolean = True)
  If vaciar Then tr.Nodes.Clear
  With busc(sql)
    Do Until .EOF
      tr.Nodes.Add IIf(.Fields(pad) = 0, Null, "k" & .Fields(pad)), tvwChild, "k" & .Fields(key), .Fields(col)
      .MoveNext
    Loop
  End With
End Sub
