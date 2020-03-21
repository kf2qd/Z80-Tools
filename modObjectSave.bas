Attribute VB_Name = "modObjectSave"
Option Explicit
Public IntelHexStr As String
Public IntelHexAddr As Long
Public IntelHexInit As Boolean
Public IntelHexCt As Integer
Public HexSel As Boolean
Public CPMSel As Boolean
Public BINSel As Boolean
Public HexInit As Boolean
Public BinInit As Boolean
Public CPMInit As Boolean
Public SrcNum As Long
Public ListNum As Long
Public ObjectNum As Long

Sub IntelHexBuild(HexAddr As Long, HexByte As String)
Dim ct As Integer
Dim chksum As Long

   If HexAddr <> IntelHexAddr Or IntelHexCt > 15 Then
      If IntelHexInit Then
         Call IntelHexFlush
      End If
      IntelHexInit = True
      IntelHexCt = 0
      IntelHexAddr = HexAddr
      IntelHexStr = ":10" + Hex4(HexAddr) + "00"
   End If
   IntelHexStr = IntelHexStr + HexByte
   IntelHexAddr = IntelHexAddr + 1
   IntelHexCt = IntelHexCt + 1

End Sub

Sub IntelHexAdd(ByVal HexAddr As Long, ByVal codestr As String)
Dim ct As Integer
Dim chksum As Long

   
   For ct = 1 To Len(codestr) Step 2
      Call IntelHexBuild(HexAddr, Mid(codestr, ct, 2))
      HexAddr = HexAddr + 1
   Next ct
      

End Sub

Sub IntelHexFlush()
Dim ct As Integer
Dim chksum As Long
   
   IntelHexStr = Left(IntelHexStr + String(32, "0"), 41)
   chksum = 0
   For ct = 2 To 40 Step 2
      chksum = chksum + val("&H" + Mid(IntelHexStr, ct, 2))
   Next ct
   chksum = chksum And &HFF
   chksum = (&H100 - chksum) And &HFF
   IntelHexStr = IntelHexStr + hex2L(chksum)
   Print #ObjectNum, IntelHexStr
   IntelHexStr = ""
   
End Sub

Sub IntelHexClose(ByVal codestr As String)
Dim ct As Integer
Dim chksum As Long
   
   Call IntelHexFlush
   chksum = 0
   If Left(codestr, 1) = "*" Then
      codestr = Right(codestr, 4)
      IntelHexStr = ":00" + codestr + "01"
      For ct = 2 To Len(IntelHexStr) Step 2
         chksum = chksum + val("&H" + Mid(IntelHexStr, ct, 2))
      Next ct
      chksum = chksum And &HFF
      chksum = (&H100 - chksum) And &HFF
      IntelHexStr = IntelHexStr + hex2L(chksum)
   Else
      IntelHexStr = ":00000001FF"
   End If
   Print #ObjectNum, IntelHexStr
   
End Sub

Sub CPMAdd(ByVal codestr As String)
Dim ct As Integer
Dim outbyte As Byte
   
   For ct = 1 To Len(codestr) Step 2
      outbyte = val("&H" + Mid(codestr, ct, 2))
      Put #3, , outbyte
   Next ct

End Sub

Sub BINAdd(ByVal codestr As String)
Dim ct As Integer
Dim outbyte As Byte
   
   For ct = 1 To Len(codestr) Step 2
      outbyte = val("&H" + Mid(codestr, ct, 2))
      Put #3, , outbyte
   Next ct

End Sub

Public Sub InitObject()
   If HexSel Then
      IntelHexStr = ""
      IntelHexInit = False
      IntelHexAddr = 65536
   ElseIf CPMSel Then
      
   ElseIf BINSel Then
      
   End If
   
End Sub

Public Sub ObjectSave(ByVal codestr As String)
Dim HexAddr As Long

   HexAddr = CLng("&H" + Left(codestr, 4))
   codestr = Right(codestr, Len(codestr) - 4)
   If HexSel Then
      
      Call IntelHexAdd(HexAddr, codestr)
   ElseIf CPMSel Then
      Call CPMAdd(codestr)
   ElseIf BINSel Then
      Call BINAdd(codestr)
   End If

End Sub

Sub ObjectClose(ByVal codestr As String)
   If HexSel Then
      Call IntelHexClose(codestr)
   ElseIf CPMSel Then
   
   ElseIf BINSel Then
'
   End If

End Sub

