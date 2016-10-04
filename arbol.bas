Attribute VB_Name = "arbol"
Option Explicit

Public Sub tildarabajo(ByVal nodo As Node)
  Dim odon As Node, i As Integer
  If nodo.Children > 0 Then
    Set odon = nodo.Child
    For i = 1 To nodo.Children
      odon.Checked = nodo.Checked
      odon.tag = nodo.Checked
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
    nodo.Parent.tag = 0
    For i = 1 To odon.Parent.Children
      nodo.Parent.Checked = nodo.Parent.Checked And odon.Checked
      nodo.Parent.tag = nodo.Parent.tag Or odon.Checked
      Set odon = odon.Next
    Next
  End If
End Sub

Public Sub llenarnivel(tr As TreeView, SQL As String, col As String, key As String, pad As String, Optional vaciar As Boolean = True)
  If vaciar Then tr.Nodes.Clear
  With exec(SQL)
    Do Until .EOF
      tr.Nodes.Add IIf(.fields(pad) = 0, Null, "k" & .fields(pad)), tvwChild, "k" & .fields(key), .fields(col)
      .MoveNext
    Loop
  End With
End Sub
