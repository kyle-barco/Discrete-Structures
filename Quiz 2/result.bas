Attribute VB_Name = "result"
Option Compare Database

Function res()
  Dim octInput As Variant
  Dim octRes As Integer
   
  octInput = Forms!Form2.txtboxOctal.Value
  
  If IsNull(Forms!Form2.txtboxOctal.Value) Or Trim(Forms!Form2.txtboxOctal.Value) = "" Then
    Forms!Form2.lblRes.Caption = "Input a Value"
    Exit Function
  End If
  
  octInput = CLng(Forms!Form2.txtboxOctal.Value)
  
  octRes = Oct(octInput)
  Forms!Form2.lblRes.Caption = octRes
  Forms!Form2.txtboxOctal.Value = ""

End Function
