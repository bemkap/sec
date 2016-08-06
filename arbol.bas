Attribute VB_Name = "arbol"
Public Function prof(nodo As Node) As Integer
  If nodo Is Nothing Then prof = 0 Else prof = IIf(nodo.Parent Is Nothing, 0, 1 + prof(nodo.Parent))
End Function

Public Sub tildarAbajo(ByVal nodo As Node)
  If nodo.Children > 0 Then
    Set odon = nodo.Child
    For i = 1 To nodo.Children
      odon.Checked = nodo.Checked
      odon.Tag = nodo.Checked
      tildarAbajo odon
      Set odon = odon.Next
    Next
  End If
End Sub

Public Sub tildarArriba(ByVal nodo As Node)
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

Public Sub llenarNivel(tr As TreeView, sql As String, col As String, key As String, pad As String, Optional vaciar As Boolean = True)
  If vaciar Then tr.Nodes.Clear
  With busc(sql)
    Do Until .EOF
      tr.Nodes.Add IIf(.Fields(pad) = 0, Null, "k" & .Fields(pad)), tvwChild, "k" & .Fields(key), .Fields(col)
      .MoveNext
    Loop
  End With
End Sub
