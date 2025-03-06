Attribute VB_Name = "result"
Option Compare Database

Function res()
  Dim binInput As String
  Dim decRes As Long
   
  binInput = Forms!Form2.txtboxBinary.Value
  
  If binInput = "" Or Not binInput Like "[01]*" Then
    Forms!Form2.lblRes.Caption = "Please Input a Binary value"
    Exit Function
  End If
  
  decRes = BinToDec(binInput)
  Forms!Form2.lblRes.Caption = decRes
  Forms!Form2.txtboxBinary.Value = ""

End Function
