Attribute VB_Name = "commovimiento"
Public Sub loadmov(frm As Form, ByRef porcs() As Double)
  porcs(0, 0) = 0.21: porcs(1, 0) = 0.105: porcs(2, 0) = 0.27
  porcs(0, 1) = 0: porcs(1, 1) = 0: porcs(2, 1) = 0
  frm.labiva = porcs(0, 0) * 100 & "%"
  llenarcmb frm.cmbletra, "select * from comprobantes", "nom_comp"
End Sub

Public Sub teclacuemov(frm As Form, KeyCode As Integer)
  On Error GoTo E
  Select Case KeyCode
  Case vbKeyF3: teclacue frm.txtcuenta, frm.txtemp, frm.labcodcue
  Case vbKeyF4: If p And 2 ^ 6 Then teclacue1 Else StatusBar1.SimpleText = "Permisos necesarios"
  End Select
  Exit Sub
E: StatusBar1.SimpleText = Err.Description
End Sub

Public Function validarcuemov(frm As Form) As Boolean
  On Error GoTo E
  If frm.txtcuenta <> "" Then validarcuemov = validarcue(frm.txtcuenta, frm.txtemp, frm.labcodcue)
  Exit Function
E: StatusBar1.SimpleText = Err.Description: frm.txtcuenta = ""
End Function

Public Sub teclatxtnmov(frm As Form, idxiva As Integer, ByRef idc As Integer, ByRef porcs() As Double)
  porcs(idc, 1) = val(frm.txtn(idxiva))
  idc = (idc + 1) Mod 3
  frm.txtn(idxiva) = Format(porcs(idc, 1), "0.00")
  frm.labiva = porcs(idc, 0) * 100 & "%"
  frm.txtn(idxiva).SelStart = 0
  frm.txtn(idxiva).SelLength = Len(frm.txtn(idxiva))
End Sub
