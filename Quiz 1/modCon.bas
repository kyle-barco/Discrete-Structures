Attribute VB_Name = "modCon"
Option Compare Database

Function BinToDec(sMyBin As String)
  Dim x As Integer
  Dim iLen As Integer
  
  iLen = Len(sMyBin) - 1
  For x = 0 To iLen
    BinToDec = BinToDec + Mid(sMyBin, iLen - x + 1, 1) * 2 ^ x
  Next x
End Function
