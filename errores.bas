Attribute VB_Name = "errores"
Enum NERROR
  NOCAMP = 1
  CAMPEX = 2
  INVOP = 3
  INVDAT = 4
End Enum

Public Sub assert(p As Boolean, n As NERROR, d As String)
  If Not p Then Err.Raise n, , d
End Sub
