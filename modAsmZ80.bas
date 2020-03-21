Attribute VB_Name = "modAsmZ80"
Option Explicit

Type SymRec
   name As String  ' Symbol name }
   value As Long      ' Symbol value }
   defined As Boolean   ' TRUE if defined }
   multiDef As Boolean  ' TRUE if multiply defined }
   equ As Boolean       ' TRUE if defined with EQU pseudo }
End Type

Const maxSymLen  As Integer = 40
Const maxOpcdLen As Integer = 4

' Const regs = "  B  C  D  E  H  L  M  A  I  R BC DE HL SP IX IY AF AF' "
'     regVals = ' 0  1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17'
Public Regs(1 To 18) As String

Const reg_None = -2
Const reg_Imed = -1
Const reg_B = 0
Const reg_C = 1
Const reg_D = 2
Const reg_E = 3
Const reg_H = 4
Const reg_L = 5
Const reg_M = 6
Const reg_A = 7
Const reg_I = 8
Const reg_R = 9
Const reg_BC = 10
Const reg_DE = 11
Const reg_HL = 12
Const reg_SP = 13
Const reg_IX = 14
Const reg_IY = 15
Const reg_AF = 16
Const reg_AFpr = 17
Const reg_Indir = 64
Const reg_Imed_Indir = reg_Imed - reg_Indir   ' -65
Const reg_C_Indir = reg_C + reg_Indir         ' 65
Const reg_BC_Indir = reg_BC + reg_Indir       ' 74
Const reg_DE_Indir = reg_DE + reg_Indir       ' 75
Const reg_HL_Indir = reg_HL + reg_Indir       ' 76
Const reg_SP_Indir = reg_SP + reg_Indir       ' 77
Const reg_IX_Indir = reg_IX + reg_Indir       ' 78
Const reg_IY_Indir = reg_IY + reg_Indir       ' 79

Public SymTabLast As Integer
Public symTable(16384) As SymRec

Public Source As Integer           ' file pointers
Public Listing As Integer
Public Object As Integer

Public Pass2 As Boolean

Public ASMPC As Long
Public PCInc As Long
Public ErrorStr As String
Public EndOp As Boolean

Dim LabelStr  As String         ' variables used during the processing of each line.
Dim OpcodeStr As String
Dim Reg1str As String
Dim Reg2str As String
Dim DispStr As String
Dim CommentStr As String

Dim codestr As String

Sub Error(errstr As String)    ' assembly errors passed here to be printed after the current line in the listing
   
   If Pass2 Then ErrorStr = ErrorStr + errstr + vbCrLf

End Sub

Public Function hex2I(I As Integer) As String   ' convert an 8 bit number to a 2 digit hex
   
   hex2I = Right("0" + Hex(I), 2)
   
End Function

Public Function hex2L(L As Long) As String   ' convert an 8 bit number to a 2 digit hex
   
   hex2L = Right("0" + Hex(L), 2)
   
End Function

Public Function Hex4(I As Long) As String
   
   Hex4 = Right("000" + Hex(I), 4)
   
End Function

Function OpcodeHex4(I As Long) As String
Dim tempstr As String
   
   tempstr = Right("000" + Hex(I), 4)
   OpcodeHex4 = Right(tempstr, 2) + Left(tempstr, 2)
   
End Function

Private Function Bin2Dec(BinStr As String) As String
Dim workStr As String
Dim WorkVal As Long
Dim ct As Integer

   workStr = StrReverse(Left(BinStr, Len(BinStr) - 1))
   For ct = 1 To Len(workStr)
      WorkVal = WorkVal + 2 ^ (ct - 1) * Val(Mid(workStr, ct, 1))
   Next ct
   Bin2Dec = WorkVal
End Function

Function Immed8(ByVal ValStr As String) As Long
Dim ct As Integer
Dim MathStr As String
Dim MathPos As Integer
Dim Vals(8) As String
Dim Ops(8) As String
Dim OpsCt As Integer
Dim hexflag As Boolean
Dim pos As Integer

   For ct = 1 To 8
      Vals(ct) = ""
      Ops(ct) = ""
   Next ct
   OpsCt = 1
   
   If Left(ValStr, 1) = "(" Then     'Check for parens and remove.
      Mid(ValStr, 1, 1) = " "
      ValStr = Left(ValStr, Len(ValStr) - 1)
      ValStr = Trim(ValStr)
   End If
   If (Left(ValStr, 2) = "$+") Or (Left(ValStr, 2) = "$-") Or (Left(ValStr, 2) = "$ ") Then
      ValStr = Hex(ASMPC) + "H" + Right(ValStr, Len(ValStr) - 1)
   End If
   While Len(ValStr) > 0
      MathPos = 0
      MathStr = ""
      ct = 2
      While (ct < Len(ValStr)) And (MathPos = 0)          ' find the position of the Math operation
         MathStr = Mid(ValStr, ct, 1)                    ' when done MathStr has the operation
         If InStr(1, "+-*/&", MathStr) Then MathPos = ct  ' and MathPos has the position in the string
         ct = ct + 1
      Wend
      If MathPos > 0 Then
         Vals(OpsCt) = Trim(Left(ValStr, MathPos - 1))
         Ops(OpsCt) = MathStr
         ValStr = Trim(Right(ValStr, Len(ValStr) - MathPos))
         OpsCt = OpsCt + 1
      Else
         Vals(OpsCt) = Trim(ValStr)
         Ops(OpsCt) = ""
         ValStr = ""
      End If
   Wend
   OpsCt = 1
   While Vals(OpsCt) <> ""
      If Left(Vals(OpsCt), 1) = "$" Then
         Mid(Vals(OpsCt), 1, 1) = " "
         Vals(OpsCt) = "&H" + Trim(Vals(OpsCt))
      ElseIf Right(Vals(OpsCt), 1) = "H" Then
         hexflag = True
         For ct = 1 To Len(Vals(OpsCt)) - 1
            pos = InStr(1, "0123456789ABCDEF", Mid(Vals(OpsCt), ct, 1))
            If pos = 0 Then hexflag = False
         Next ct
         If hexflag Then
            Vals(OpsCt) = "&H" + Left(Vals(OpsCt), Len(Vals(OpsCt)) - 1)
         End If
      ElseIf Right(Vals(OpsCt), 1) = "B" Then
         hexflag = True
         For ct = 1 To Len(Vals(OpsCt)) - 1
            pos = InStr(1, "01", Mid(Vals(OpsCt), ct, 1))
            If pos = 0 Then hexflag = False
         Next ct
         If hexflag Then
            Vals(OpsCt) = Bin2Dec(Vals(OpsCt))
         End If
      End If
      If Val(Vals(OpsCt)) <> 0 Then
         Vals(OpsCt) = "&H" + Hex(CLng(Vals(OpsCt)) And CLng("&HFFFF")) ' 65535 = &HFFF, if &HFFFF ia used it = -1 integer
      ElseIf Vals(OpsCt) = "0" Then
         Vals(OpsCt) = 0
      ElseIf Left(Vals(OpsCt), 3) = "&H0" Then
         Vals(OpsCt) = "&H" + Hex(CLng(Vals(OpsCt)) And CLng("&HFFFF"))
      Else
         Vals(OpsCt) = "&H" + Hex(FindLabel(Vals(OpsCt)) And CLng("&HFFFF"))
      End If
      OpsCt = OpsCt + 1
   Wend
   
   Immed8 = CLng(Vals(1))
   OpsCt = 1
   While Ops(OpsCt) <> ""
      Select Case Ops(OpsCt)
         Case "+": Immed8 = Immed8 + CLng(Vals(OpsCt + 1))
         Case "-": Immed8 = Immed8 - CLng(Vals(OpsCt + 1))
         Case "*": Immed8 = Immed8 * CLng(Vals(OpsCt + 1))
         Case "/": Immed8 = Immed8 / CLng(Vals(OpsCt + 1))
         Case "&": Immed8 = Immed8 And CLng(Vals(OpsCt + 1))
      End Select
      OpsCt = OpsCt + 1
   Wend
   Immed8 = Immed8 And CLng(255)

End Function

Function Immed16(ByVal ValStr As String) As Long
Dim ct As Integer
Dim MathStr As String
Dim MathPos As Integer
Dim Vals(8) As String
Dim Ops(8) As String
Dim OpsCt As Integer
Dim hexflag As Boolean
Dim pos As Integer
Dim workStr As String


   For ct = 1 To 8
      Vals(ct) = ""
      Ops(ct) = ""
   Next ct
   OpsCt = 1
   If ValStr = "$" Then ValStr = "&H" + Hex4(ASMPC)
   If Left(ValStr, 1) = "(" Then     'Check for parens and remove.
      Mid(ValStr, 1, 1) = " "
      ValStr = Left(ValStr, Len(ValStr) - 1)
      ValStr = Trim(ValStr)
   End If
   If (Left(ValStr, 2) = "$+") Or (Left(ValStr, 2) = "$-") Or (Left(ValStr, 2) = "$ ") Then
      ValStr = "&H" + Hex(ASMPC) + Right(ValStr, Len(ValStr) - 1)
   End If
   While Len(ValStr) > 0
      MathPos = 0
      MathStr = ""
      ct = 2
      While (ct < Len(ValStr)) And (MathPos = 0)          ' find the position of the Math operation
         MathStr = Mid(ValStr, ct, 1)                    ' when done MathStr has the operation
         If InStr(1, "+-*/&", MathStr) Then MathPos = ct  ' and MathPos has the position in the string
         ct = ct + 1
      Wend
      If MathPos > 0 Then
         Vals(OpsCt) = Trim(Left(ValStr, MathPos - 1))
         Ops(OpsCt) = MathStr
         ValStr = Trim(Right(ValStr, Len(ValStr) - MathPos))
         OpsCt = OpsCt + 1
      Else
         Vals(OpsCt) = ValStr
         Ops(OpsCt) = ""
         ValStr = ""
      End If
   Wend
   OpsCt = 1
   While Vals(OpsCt) <> ""
      If Left(Vals(OpsCt), 1) = "$" Then
         Mid(Vals(OpsCt), 1, 1) = " "
         Vals(OpsCt) = "&H" + Trim(Vals(OpsCt))
      ElseIf Right(Vals(OpsCt), 1) = "H" Then
         hexflag = True
         For ct = 1 To Len(Vals(OpsCt)) - 1
            pos = InStr(1, "0123456789ABCDEF", Mid(Vals(OpsCt), ct, 1))
            If pos = 0 Then hexflag = False
         Next ct
         If hexflag Then
            Vals(OpsCt) = "&H0" + Left(Vals(OpsCt), Len(Vals(OpsCt)) - 1)
         End If
      ElseIf Right(Vals(OpsCt), 1) = "B" Then
         hexflag = True
         For ct = 1 To Len(Vals(OpsCt)) - 1
            pos = InStr(1, "01", Mid(Vals(OpsCt), ct, 1))
            If pos = 0 Then hexflag = False
         Next ct
         If hexflag Then
            Vals(OpsCt) = Bin2Dec(Vals(OpsCt))
         End If
      ElseIf Left(Vals(OpsCt), 2) = "&0" Then
         Vals(OpsCt) = "&H" + Left(Vals(OpsCt), Len(Vals(OpsCt)) - 1)
      End If
      If Val(Vals(OpsCt)) <> 0 Then
         If Left(Vals(OpsCt), 2) <> "&H" Then
            On Error GoTo Immed16Error1
            workStr = CLng(Vals(OpsCt))
            Vals(OpsCt) = "&H" + Hex(CLng(Vals(OpsCt)))
Immed16Error1:
            GoTo SkipError1
            Call Error("### Bad Number - " + Vals(OpsCt))
            GoTo SkipError1
            Vals(OpsCt) = "&H0"
            Resume Next
SkipError1:
         End If
      ElseIf Vals(OpsCt) = "0" Then
         Vals(OpsCt) = 0
      ElseIf Left(Vals(OpsCt), 3) = "&H0" Then
         Vals(OpsCt) = "&H" + Hex(CLng(Vals(OpsCt)) And CLng("&HFFFF"))
      Else
         On Error GoTo Immed16Error2
         Vals(OpsCt) = "&H" + Hex(FindLabel(Vals(OpsCt)) And CLng("&HFFFF"))
Immed16Error2:
         GoTo SkipError2
         Call Error("### Bad Number - " + Vals(OpsCt))
         Vals(OpsCt) = "&H0"
         Resume Next
SkipError2:
      End If
      OpsCt = OpsCt + 1
   Wend
   
   Immed16 = CLng(Vals(1))
   OpsCt = 1
   While Ops(OpsCt) <> ""
      Select Case Ops(OpsCt)
         Case "+": Immed16 = Immed16 + CLng(Vals(OpsCt + 1))
         Case "-": Immed16 = Immed16 - CLng(Vals(OpsCt + 1))
         Case "*": Immed16 = Immed16 * CLng(Vals(OpsCt + 1))
         Case "/": Immed16 = Immed16 / CLng(Vals(OpsCt + 1))
         Case "&": Immed16 = Immed16 And CLng(Vals(OpsCt + 1))
      End Select
      OpsCt = OpsCt + 1
   Wend
   Immed16 = Immed16 And CLng("&HFFFF")
End Function

Function PrepLine(SrcLine As String)
Dim ct As Long
   
   If Len(SrcLine) > 0 Then
      'SrcLine = UCase(SrcLine)
      For ct = 1 To Len(SrcLine)
         If Mid(SrcLine, ct, 1) = Chr(9) Then Mid(SrcLine, ct, 1) = " " ' replace tabs with spaces, easier to parse...
      Next ct
      PrepLine = SrcLine
   End If
   
End Function

Private Function UpCase(SrcStr As String) As String
Dim ct As Integer
Dim QtFlg As Boolean

   QtFlg = False
   For ct = 1 To Len(SrcStr)
      If (Mid(SrcStr, ct, 1) = "'") Or (Mid(SrcStr, ct, 1) = Chr(34)) Then
      End If
   Next ct
End Function

Sub Parse(ByVal ParseLine As String)
Dim ct As Long
Dim SrchChar As String
Dim TokenEnd As String
Dim tokenFound As Boolean
Dim Quoted As Boolean
Dim LastQuote As String
   LabelStr = ""
   OpcodeStr = ""
   Reg1str = ""
   Reg2str = ""
   CommentStr = ""
   
      ParseLine = RTrim(ParseLine)
   If Len(ParseLine) = 0 Then Exit Sub
   If Left(ParseLine, 1) = ";" Then Exit Sub
   If Left(ParseLine, 1) = " " Then   ' Check for no label at the start of the line
      LabelStr = ""
      ParseLine = Trim(ParseLine)
   Else                               ' Process out the label - if found
      ct = 1
      While (Mid(ParseLine, ct, 1) <> " ") And (Mid(ParseLine, ct, 1) <> ":") And (ct < Len(ParseLine))
         ct = ct + 1
      Wend
      LabelStr = UCase(Left(ParseLine, ct))
      If Right(LabelStr, 1) = ":" Then LabelStr = Left(LabelStr, Len(LabelStr) - 1)
      ParseLine = Trim(Right(ParseLine, Len(ParseLine) - ct))
   End If
   If (Left(ParseLine, 1) <> ";") Then        ' Opcode or comment?
      ct = 1
      While (Mid(ParseLine, ct, 1) <> " ") And (Mid(ParseLine, ct, 1) <> ";") And (ct < Len(ParseLine))
         ct = ct + 1
      Wend
      OpcodeStr = Trim(UCase(Left(ParseLine, ct)))
      If Left(OpcodeStr, 1) = "." Then OpcodeStr = Right(OpcodeStr, Len(OpcodeStr) - 1)
      If Len(ParseLine) > 0 Then
         ParseLine = Trim(Right(ParseLine, Len(ParseLine) - ct))
      End If
   End If
   If Len(ParseLine) = 0 Then Exit Sub
   If (Left(ParseLine, 1) = ";") Then Exit Sub
   If (OpcodeStr = "DEFB") Or (OpcodeStr = "DEFM") Or (OpcodeStr = "BYTE") Or (OpcodeStr = "TEXT") Then
      OpcodeStr = "DB"
   End If
   If (OpcodeStr = "DEFW") Or (OpcodeStr = "WORD") Then
      OpcodeStr = "DW"
   End If
   If (OpcodeStr = "DB") Or (OpcodeStr = "DW") Or (OpcodeStr = "EQU") Then
      TokenEnd = ";"
   Else
      TokenEnd = " ,;"
   End If
   ParseLine = ParseLine + " "
   ct = 1
   tokenFound = False
   Quoted = False
   SrchChar = ""
   LastQuote = ""
   While Not tokenFound
      SrchChar = Mid(ParseLine, ct, 1)
      If LastQuote = "" Then
         If (SrchChar = "'") Or (SrchChar = Chr(34)) Then
            Quoted = True
            LastQuote = SrchChar
         End If
      ElseIf (LastQuote = "'") And (SrchChar = "'") Then
         Quoted = False
         LastQuote = ""
      ElseIf (LastQuote = Chr(34)) And (SrchChar = Chr(34)) Then
         Quoted = False
         LastQuote = ""
      End If
      If Not Quoted Then
         If ct > Len(ParseLine) Then
            ct = Len(ParseLine)
            tokenFound = True
         Else
            SrchChar = Mid(ParseLine, ct, 1)
            Mid(ParseLine, ct, 1) = UCase(Mid(ParseLine, ct, 1))
            If InStr(1, TokenEnd, SrchChar) > 0 Then tokenFound = True
         End If
      End If
      If Not tokenFound Then ct = ct + 1
   Wend
   ct = ct - 1
   Reg1str = RTrim(Left(ParseLine, ct))
   ParseLine = Right(ParseLine, Len(ParseLine) - ct)
   ParseLine = Trim(ParseLine)
   
   If (ParseLine = "") Then Exit Sub
   If (Left(ParseLine, 1) = ";") Then Exit Sub
   If Left(ParseLine, 1) = "," Then ParseLine = Right(ParseLine, Len(ParseLine) - 1)
   If Left(ParseLine, 3) = "AF'" Then
      Reg2str = Left(ParseLine, 3)
      Exit Sub
   End If
   ParseLine = Trim(ParseLine)
   ParseLine = ParseLine + " "
   ct = 1
   tokenFound = False
   Quoted = False
   SrchChar = ""
   LastQuote = ""
   While Not tokenFound
      SrchChar = Mid(ParseLine, ct, 1)
      If LastQuote = "" Then
         If (SrchChar = "'") Or (SrchChar = Chr(34)) Then
            Quoted = True
            LastQuote = SrchChar
         End If
      ElseIf (LastQuote = "'") And (SrchChar = "'") Then
         Quoted = False
         LastQuote = ""
      ElseIf (LastQuote = Chr(34)) And (SrchChar = Chr(34)) Then
         Quoted = False
         LastQuote = ""
      End If
      If Not Quoted Then
         If ct > Len(ParseLine) Then
            ct = Len(ParseLine)
            tokenFound = True
         Else
            SrchChar = Mid(ParseLine, ct, 1)
            Mid(ParseLine, ct, 1) = UCase(Mid(ParseLine, ct, 1))
            If InStr(1, TokenEnd, SrchChar) > 0 Then tokenFound = True
         End If
      End If
      If Not tokenFound Then ct = ct + 1
   Wend
   ct = ct - 1
   Reg2str = Trim(Left(ParseLine, ct))
   
End Sub

Sub AddLabel(Label As String, address As Long)
Dim ct As Long
Dim Duplicate As Boolean
   
   If Label = "" Then Exit Sub
   For ct = 1 To SymTabLast
      If symTable(ct).name = Label Then
         Duplicate = True
         symTable(ct).multiDef = True
      End If
   Next ct
   If Not Duplicate Then
      symTable(SymTabLast).name = Trim(Label)
      symTable(SymTabLast).value = address
      symTable(SymTabLast).defined = True
      SymTabLast = SymTabLast + 1
   End If
End Sub

Function FindLabel(LblStr As String) As Long
Dim ct As Long
Dim SymFound As Boolean
   
   SymFound = False
   ct = 1
   While (ct < SymTabLast) And (Not SymFound)
      If symTable(ct).name = LblStr Then
         SymFound = True
      Else
         ct = ct + 1
      End If
   Wend
   If SymFound Then
      FindLabel = symTable(ct).value
   Else
      FindLabel = 0
      If Pass2 Then Call Error("### Undefined Symbol Referenced - " + LblStr)
   End If
   
End Function

Function RegTyp(ByVal RegStr As String) As Long
'Regs   B   C   D   E   H   L   (HL) A   I   R   BC  DE  HL  SP  IX  IY  AF  AF'
'       0   1   2   3   4   5   6    7   8   9   10  11  12  13  14  15  16  17
Dim RegNo As Long
Dim Indir As Long
Dim RegFound As Boolean
Dim RegTypWOrk

   
   If Left(RegStr, 1) = "(" Then
      Mid(RegStr, 1) = " "
      RegStr = Left(RegStr, Len(RegStr) - 1)
      RegStr = Trim(RegStr)
      Indir = reg_Indir
   Else
      Indir = 0
   End If
   RegFound = False
   RegNo = 1
   If Left(RegStr, 2) = "IX" Then RegStr = "IX"
   If Left(RegStr, 2) = "IY" Then RegStr = "IY"
   While (Not RegFound) And (RegNo < reg_AFpr + 2)
      If RegStr = Regs(RegNo) Then
         RegFound = True
      Else
         RegNo = RegNo + 1
      End If
   Wend
   If RegFound Then
      RegTypWOrk = (RegNo - 1) + Indir
   ElseIf RegStr = "" Then
      RegTypWOrk = reg_None
   Else
      RegTypWOrk = reg_Imed - Indir
   End If
   If RegTypWOrk = reg_HL_Indir Then RegTypWOrk = reg_M
         
   RegTyp = RegTypWOrk

End Function

Function CalcRelOffset(PC As Long, Destination As String) As String
Dim DestTmp As Long
Dim OfsPosFLG As Boolean
Dim ct As Integer

   If Val(Destination) <> 0 Then
      DestTmp = Immed16(Destination)
   End If
   OfsPosFLG = False
   If Left(Destination, 1) = "$" Then
      ct = 1
      While (ct < Len(Destination)) And (Mid(Destination, ct, 1) <> "+") And (Mid(Destination, ct, 1) <> "-")
         ct = ct + 1
      Wend
      If Mid(Destination, ct, 1) = "+" Then OfsPosFLG = True
      Mid(Destination, 1, ct) = "                         "
      If OfsPosFLG Then
         CalcRelOffset = hex2L(Val(Destination And &H7F))
      Else
         CalcRelOffset = hex2L((&H100 - Val(Destination)) And &HFF)
      End If
   ElseIf DestTmp <> 0 Then
      CalcRelOffset = hex2L((DestTmp - PC) And &HFF)
   Else
      CalcRelOffset = hex2L((FindLabel(Destination) - PC) And &HFF)
   End If
   
End Function

Function XYDisplacement(ByVal RegStr As String) As String
Dim CutPos As Integer
   CutPos = InStr(1, RegStr, "+")
   If CutPos = 0 Then
      XYDisplacement = "00"
   Else
      Mid(RegStr, 1, CutPos) = "          "
      RegStr = Trim(RegStr)
      RegStr = Left(RegStr, Len(RegStr) - 1)
      XYDisplacement = hex2L(Immed8(RegStr))
   End If
   
End Function

Private Function Bytes(ByteStr As String) As String
Dim CommaPos As Integer
Dim WorkStr1 As String
Dim workstr2 As String
Dim HoldStr As String
Dim ct As Integer

   HoldStr = ""
   While Len(ByteStr) > 0
      If ((Left(ByteStr, 1) = "'") And (Mid(ByteStr, 3, 1) = "'")) Or ((Left(ByteStr, 1) = Chr(34)) And (Mid(ByteStr, 3, 1) = Chr(34))) Then
         CommaPos = InStr(3, ByteStr, ",")
      Else
         CommaPos = InStr(1, ByteStr, ",")
      End If
      If CommaPos > 0 Then
         WorkStr1 = Trim(Left(ByteStr, CommaPos - 1))
         ByteStr = Right(ByteStr, Len(ByteStr) - CommaPos)
      Else
         WorkStr1 = ByteStr
         ByteStr = ""
      End If
      If (Left(WorkStr1, 1) = "'") And (Mid(WorkStr1, 3, 1) = "'") Then
          workstr2 = Mid(WorkStr1, 2, 1)
          workstr2 = "&H" + hex2L(Asc(workstr2)) + Right(WorkStr1, Len(WorkStr1) - 3)
          WorkStr1 = workstr2
      End If
      If (Left(WorkStr1, 1) = Chr(34)) And (Mid(WorkStr1, 3, 1) = Chr(34)) Then
          workstr2 = Mid(WorkStr1, 2, 1)
          workstr2 = "&H" + hex2L(Asc(workstr2)) + Right(WorkStr1, Len(WorkStr1) - 3)
          WorkStr1 = workstr2
      End If
      If ((Left(WorkStr1, 1) = "'" And Right(WorkStr1, 1) = "'")) Or ((Left(WorkStr1, 1) = Chr(34)) And (Right(WorkStr1, 1) = Chr(34))) Then
         Mid(WorkStr1, 1, 1) = " "
         Mid(WorkStr1, Len(WorkStr1), 1) = " "
         WorkStr1 = Trim(WorkStr1)
         HoldStr = HoldStr + WorkStr1
      Else
         HoldStr = HoldStr + Chr(Immed8(WorkStr1))
      End If
   Wend
   For ct = 1 To Len(HoldStr)
      Bytes = Bytes + hex2L(Asc(Mid(HoldStr, ct, 1)))
   Next ct
End Function

Private Function Words(WordStr As String) As String
Dim CommaPos As Integer
Dim WorkStr1 As String
Dim workstr2 As String
Dim HoldStr As String
Dim retval As Long
Dim RetValHi As Long
Dim RetvalLo As Long
Dim ct As Integer

   HoldStr = ""
   While Len(WordStr) > 0
      CommaPos = InStr(1, WordStr, ",")
      If CommaPos > 0 Then
         WorkStr1 = Trim(Left(WordStr, CommaPos - 1))
         WordStr = Right(WordStr, Len(WordStr) - CommaPos)
      Else
         WorkStr1 = WordStr
         WordStr = ""
      End If
      retval = Immed16(WorkStr1)
      RetValHi = (retval \ 256) And &HFF
      RetvalLo = retval And &HFF
      HoldStr = HoldStr + Chr(RetvalLo) + Chr(RetValHi)
   Wend
   For ct = 1 To Len(HoldStr)
      Words = Words + hex2L(Asc(Mid(HoldStr, ct, 1)))
   Next ct
End Function

' the subroutine InitVars() clears teh symbol table and initializes the Regs() array.
' should be called at the start of assembly of a file, and if symbols will be used  when
' assembling one line at a time in the emulator.

Public Sub InitVars()
Dim ct As Integer

   SymTabLast = 1
   For ct = 1 To 16384
      symTable(ct).defined = False
      symTable(ct).equ = False
      symTable(ct).multiDef = False
      symTable(ct).name = "##########"
      symTable(ct).value = 0
   Next ct
   
   Regs(1) = "B"
   Regs(2) = "C"
   Regs(3) = "D"
   Regs(4) = "E"
   Regs(5) = "H"
   Regs(6) = "L"
   Regs(7) = "(HL)"
   Regs(8) = "A"
   Regs(9) = "I"
   Regs(10) = "R"
   Regs(11) = "BC"
   Regs(12) = "DE"
   Regs(13) = "HL"
   Regs(14) = "SP"
   Regs(15) = "IX"
   Regs(16) = "IY"
   Regs(17) = "AF"
   Regs(18) = "AF'"

End Sub

' this function will return the address and object code for the line of assembly code passed
' the returned string will have 4 hex digits for the for the start address for the object code
' and pairs of hex digits representing the object code will follow.
' Prior to calling this function the first time the subroutine InitVars() should be called 1 time.


Public Function Assemble(ASMLine As String) As String
Dim Reg1Typ As Long
Dim Reg2Typ As Long

Dim CodeTemp As Long
Dim SrcLine As String
Dim RelOffset As Long
Dim tmpInt As Integer
Dim tmpLong As Long
Dim tmpStr As String
Dim tmpstr2 As String
   

   
   codestr = ""
   ASMPC = ASMPC + PCInc
   PCInc = 0
   SrcLine = PrepLine(ASMLine)
   Call Parse(SrcLine)
   Reg1Typ = RegTyp(Reg1str)
   Reg2Typ = RegTyp(Reg2str)
   
   If (Not Pass2) And (OpcodeStr <> "EQU") And (LabelStr <> "") Then
      Call AddLabel(LabelStr, ASMPC)
   End If
   Select Case OpcodeStr
      Case "":   'blank line
      Case "=", "EQU":
                       tmpLong = Immed16(Reg1str)
                       codestr = "= " + Hex4(tmpLong)
                       If (Not Pass2) And (LabelStr <> "") Then
                          Call AddLabel(LabelStr, tmpLong)
                       End If
      Case "DS", "DEFS":
                         tmpLong = Immed16(Reg1str)
                         codestr = String(tmpLong * 2, "0")
                         PCInc = tmpLong
      Case "DW", "DEFW", "WORD":
                         codestr = Words(Reg1str)
                         PCInc = Len(codestr) / 2
      Case "DB", "DEFB", "DEFM", "BYTE", "TEXT":
                         codestr = Bytes(Reg1str)
                         PCInc = Len(codestr) / 2
      Case "ORG":
                  ASMPC = Immed16(Reg1str)
      Case "END":
                  If Pass2 Then
                     EndOp = True
                     If Reg1str <> "" Then
                        codestr = "*" + Hex4(Immed16(Reg1str))
                     Else
                        codestr = "-"
                     End If
                  End If
      Case "NOP": codestr = "00"
                  PCInc = 1
      Case "RLCA": codestr = "07"
                  PCInc = 1
      Case "RRCA": codestr = "0F"
                  PCInc = 1
      Case "RLA": codestr = "17"
                  PCInc = 1
      Case "RRA": codestr = "1F"
                  PCInc = 1
      Case "DAA": codestr = "27"
                  PCInc = 1
      Case "CPL": codestr = "2F"
                  PCInc = 1
      Case "SCF": codestr = "37"
                  PCInc = 1
      Case "CCF": codestr = "3F"
                  PCInc = 1
      Case "HALT": codestr = "76"
                  PCInc = 1
      Case "EXX": codestr = "D9"
                  PCInc = 1
      Case "DI": codestr = "F3"
                  PCInc = 1
      Case "EI": codestr = "FB"
                  PCInc = 1
      Case "NEG": codestr = "ED44"
                  PCInc = 2
      Case "RETN": codestr = "ED45"
                  PCInc = 2
      Case "RETI": codestr = "ED4D"
                  PCInc = 2
      Case "RRD": codestr = "ED67"
                  PCInc = 2
      Case "RLD": codestr = "ED6F"
                  PCInc = 2
      Case "LDI": codestr = "EDA0"
                  PCInc = 2
      Case "CPI": codestr = "EDA1"
                  PCInc = 2
      Case "INI": codestr = "EDA2"
                  PCInc = 2
      Case "OUTI": codestr = "EDA3"
                  PCInc = 2
      Case "LDD": codestr = "EDA8"
                  PCInc = 2
      Case "CPD": codestr = "EDA9"
                  PCInc = 2
      Case "IND": codestr = "EDAA"
                  PCInc = 2
      Case "OUTD": codestr = "EDAB"
                  PCInc = 2
      Case "LDIR": codestr = "EDB0"
                  PCInc = 2
      Case "CPIR": codestr = "EDB1"
                  PCInc = 2
      Case "INIR": codestr = "EDB2"
                  PCInc = 2
      Case "OTIR": codestr = "EDB3"
                  PCInc = 2
      Case "LDDR": codestr = "EDB8"
                  PCInc = 2
      Case "CPDR": codestr = "EDB9"
                  PCInc = 2
      Case "INDR": codestr = "EDBA"
                  PCInc = 2
      Case "OTDR": codestr = "EDBB"
                  PCInc = 2
      Case "EX": PCInc = 1
                 If Reg1Typ = reg_AF Then
                    If Reg2Typ = reg_AFpr Then
                       codestr = "08"
                    Else: Call Error("### Bad Source Operand - " + Reg2str)
                    End If
                 ElseIf (Reg1Typ = reg_SP + reg_Indir) Then
                    Select Case Reg2Typ
                       Case reg_HL: codestr = "E3"
                       Case reg_IX: codestr = "DDE3"
                                    PCInc = PCInc + 1
                       Case reg_IY: codestr = "FDE3"
                                    PCInc = PCInc + 1
                       Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                    End Select
                 ElseIf (Reg1Typ = reg_DE) Then
                    If Reg2Typ = reg_HL Then
                       codestr = "EB"
                    Else: Call Error("### Bad Source Operand -" + Reg2str)
                    End If
                 Else: Call Error("### Bad Destination Operand - " + Reg1str)
                 End If   ' EX
      Case "LD": Select Case Reg1Typ
                    Case reg_A:  PCInc = 1
                                 Select Case Reg2Typ
                                    Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                                       codestr = hex2L(&H78 + Reg2Typ)
                                    Case reg_Imed
                                       codestr = "3E" + Bytes(Reg2str)
                                       PCInc = PCInc + 1
                                    Case reg_I
                                       codestr = "ED57"
                                       PCInc = PCInc + 1
                                    Case reg_R
                                       codestr = "5F"
                                    Case reg_BC_Indir
                                       codestr = "0A"
                                    Case reg_DE_Indir
                                       codestr = "1A"
                                    Case reg_IX_Indir
                                       codestr = "DD7E" + XYDisplacement(Reg2str)
                                       PCInc = PCInc + 2
                                    Case reg_IY_Indir
                                       codestr = "FD7E" + XYDisplacement(Reg2str)
                                       PCInc = PCInc + 2
                                    Case reg_Imed_Indir
                                       codestr = "3A" + OpcodeHex4(Immed16(Reg2str))
                                       PCInc = PCInc + 2
                                    Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                                End Select
                    Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M:
                                Select Case Reg2Typ
                                   Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                                      codestr = hex2L(&H40 + (Reg1Typ * 8) + Reg2Typ)
                                      PCInc = 1
                                   Case reg_Imed
                                      codestr = hex2L(6 + Reg1Typ * 8) + Bytes(Reg2str)
                                      PCInc = 2
                                   Case reg_IX_Indir
                                      PCInc = 3
                                      codestr = "DD" + hex2L(&H46 + Reg1Typ * 8) + XYDisplacement(Reg2str)
                                   Case reg_IY_Indir
                                      PCInc = 3
                                      codestr = "FD" + hex2L(&H46 + Reg1Typ * 8) + XYDisplacement(Reg2str)
                                   Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                                End Select
                    Case reg_I:
                                If Reg2Typ = reg_A Then
                                   codestr = "ED47"
                                   PCInc = 2
                                 Else
                                    Call Error("### Bad Source Operand - " + Reg2str)
                                 End If
                    Case reg_R:
                                If Reg2Typ = reg_A Then
                                   codestr = "ED4F"
                                   PCInc = 2
                                 Else
                                    Call Error("### Bad Source Operand - " + Reg2str)
                                 End If
                    Case reg_BC, reg_DE, reg_HL, reg_SP, reg_IX, reg_IY:
                                 If Reg1Typ = reg_IX Then
                                    codestr = "DD"
                                    PCInc = PCInc + 1
                                    Reg1Typ = reg_HL
                                 ElseIf Reg1Typ = reg_IY Then
                                    codestr = "FD"
                                    PCInc = PCInc + 1
                                    Reg1Typ = reg_HL
                                 End If
                                 If Reg2Typ = reg_Imed Then
                                    codestr = codestr + hex2L((Reg1Typ - 10) * 16 + 1)
                                    codestr = codestr + OpcodeHex4(Immed16(Reg2str))
                                    PCInc = PCInc + 3
                                 ElseIf (Reg2Typ = reg_Imed_Indir) Then
                                    Select Case Reg1Typ
                                       Case reg_BC: codestr = "ED4B" + OpcodeHex4(Immed16(Reg2str))
                                                    PCInc = 4
                                       Case reg_DE: codestr = "ED5B" + OpcodeHex4(Immed16(Reg2str))
                                                    PCInc = 4
                                       Case reg_HL: codestr = codestr + "2A" + OpcodeHex4(Immed16(Reg2str))
                                                    PCInc = PCInc + 3
                                       Case reg_SP: codestr = "ED7B" + OpcodeHex4(Immed16(Reg2str))
                                                    PCInc = 4
                                       Case Else: Call Error("### Bad Destination Register - " + Reg1str)
                                    End Select
                                 ElseIf (Reg1Typ = reg_SP) Then
                                    Select Case Reg2Typ
                                       Case reg_HL: codestr = "F9"
                                                    PCInc = 1
                                       Case reg_IX: codestr = "DDF9"
                                                    PCInc = 2
                                       Case reg_IY: codestr = "FDF9"
                                                    PCInc = 2
                                       Case Else: Error ("### Bad source Operand - " + Reg2str)
                                    End Select
                                 Else
                                    Call Error("### Bad Operand - " + Reg1str)
                                 End If
                    Case reg_BC_Indir:
                                 If Reg2Typ = reg_A Then
                                    codestr = "02"
                                    PCInc = 1
                                 Else
                                    Call Error("### Bad Source Operand - " + Reg2str)
                                 End If
                    Case reg_DE_Indir:
                                 If Reg2Typ = reg_A Then
                                    codestr = "12"
                                    PCInc = 1
                                 Else
                                    Call Error("### Bad Source Operand - " + Reg2str)
                                 End If
                    Case reg_Imed_Indir:
                                 Select Case Reg2Typ
                                    Case reg_BC: codestr = "ED43" + OpcodeHex4(Immed16(Reg1str))
                                                 PCInc = 4
                                    Case reg_DE: codestr = "ED53" + OpcodeHex4(Immed16(Reg1str))
                                                 PCInc = 4
                                    Case reg_HL: codestr = "22" + OpcodeHex4(Immed16(Reg1str))
                                                 PCInc = 3
                                    Case reg_SP: codestr = "ED73" + OpcodeHex4(Immed16(Reg1str))
                                                 PCInc = 4
                                    Case reg_A:
                                                codestr = "32" + OpcodeHex4(Immed16(Reg1str))
                                                PCInc = 3
                                    Case reg_IX: codestr = "DD22" + OpcodeHex4(Immed16(Reg1str))
                                                 PCInc = 4
                                    Case reg_IY: codestr = "FD22" + OpcodeHex4(Immed16(Reg1str))
                                                 PCInc = 4
                                    Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                                 End Select
                    Case reg_IX_Indir:
                                 Select Case Reg2Typ
                                    Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_A:
                                       PCInc = 3
                                       codestr = "DD" + hex2L(&H70 + Reg2Typ) + XYDisplacement(Reg1str)
                                    Case reg_Imed:
                                       If Reg1str = "(IX)" Then
                                          PCInc = 4
                                          codestr = "DD2A" + OpcodeHex4(Immed16(Reg2str))
                                       Else
                                          PCInc = 4
                                          codestr = "DD36" + XYDisplacement(Reg1str) + hex2L(Immed8(Reg2str))
                                       End If
                                    Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                                 End Select
                    Case reg_IY_Indir:
                                 Select Case Reg2Typ
                                    Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_A:
                                       PCInc = 3
                                       codestr = "FD" + hex2L(&H70 + Reg2Typ) + XYDisplacement(Reg1str)
                                    Case reg_Imed:
                                       If Reg1str = "(IX)" Then
                                          PCInc = 4
                                          codestr = "FD2A" + OpcodeHex4(Immed16(Reg2str))
                                       Else
                                          PCInc = 4
                                          codestr = "FD36" + XYDisplacement(Reg1str) + hex2L(Immed8(Reg2str))
                                       End If
                                    Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                                 End Select
                    Case Else:  Call Error("### Bad Destination Operand - " + Reg1str)
                 End Select  ' LD
      Case "INC": PCInc = 1
                  Select Case Reg1Typ
                    Case reg_B: codestr = "04"
                    Case reg_C: codestr = "0C"
                    Case reg_D: codestr = "14"
                    Case reg_E: codestr = "1C"
                    Case reg_H: codestr = "24"
                    Case reg_L: codestr = "2C"
                    Case reg_M: codestr = "34"
                    Case reg_A: codestr = "3C"
                    Case reg_BC: codestr = "03"
                    Case reg_DE: codestr = "13"
                    Case reg_HL: codestr = "23"
                    Case reg_SP: codestr = "33"
                    Case reg_IX: codestr = "DD23"
                                 PCInc = PCInc + 1
                    Case reg_IY: codestr = "FD23"
                                 PCInc = PCInc + 1
                    Case reg_IX_Indir:
                                 codestr = "DD34" + XYDisplacement(Reg1str)
                                 PCInc = PCInc + 2
                     Case reg_IY_Indir:
                                 codestr = "FD34" + XYDisplacement(Reg1str)
                                 PCInc = PCInc + 2
                    Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                  End Select  ' INC
      Case "DEC": Select Case Reg1Typ
                    Case reg_B: codestr = "05"
                                PCInc = 1
                    Case reg_C: codestr = "0D"
                                PCInc = 1
                    Case reg_D: codestr = "15"
                                PCInc = 1
                    Case reg_E: codestr = "1D"
                                PCInc = 1
                    Case reg_H: codestr = "25"
                                PCInc = 1
                    Case reg_L: codestr = "2D"
                                PCInc = 1
                    Case reg_M: codestr = "35"
                                PCInc = 1
                    Case reg_A: codestr = "3D"
                                PCInc = 1
                    Case reg_BC: codestr = "0B"
                                 PCInc = 1
                    Case reg_DE: codestr = "1B"
                                 PCInc = 1
                    Case reg_HL: codestr = "2B"
                                 PCInc = 1
                    Case reg_SP: codestr = "3B"
                                 PCInc = 1
                    Case reg_IX: codestr = "DD2B"
                                 PCInc = 2
                    Case reg_IY: codestr = "FD2B"
                                 PCInc = 2
                    Case reg_IX_Indir:
                                 codestr = "DD35" + XYDisplacement(Reg1str)
                                 PCInc = 3
                     Case reg_IY_Indir:
                                 codestr = "FD35" + XYDisplacement(Reg1str)
                                 PCInc = 3
                    Case Else: Call Error("### Bad Destination Operand - " + Reg1str)
                  End Select  ' DEC
      Case "ADD": Select Case Reg1Typ
                     Case reg_A:
                        Select Case Reg2Typ
                           Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                              codestr = Hex(&H80 + Reg2Typ)
                              PCInc = 1
                           Case reg_Imed
                              codestr = "C6" + Bytes(Reg2str)
                              PCInc = 2
                           Case reg_IX_Indir:
                              codestr = "DD86" + XYDisplacement(Reg2str)
                              PCInc = 3
                           Case reg_IY_Indir:
                              codestr = "FD86" + XYDisplacement(Reg2str)
                              PCInc = 3
                           Case Else:  Call Error("### Bad Source Operand - " + Reg2str)
                        End Select
                     Case reg_HL: Select Case Reg2Typ
                                     Case reg_BC: codestr = "09"
                                                     PCInc = 1
                                     Case reg_DE: codestr = "19"
                                                     PCInc = 1
                                     Case reg_HL: codestr = "29"
                                                     PCInc = 1
                                     Case reg_SP: codestr = "39"
                                                     PCInc = 1
                                     Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                                  End Select
                     Case reg_IX: Select Case Reg2Typ
                                     Case reg_BC: codestr = "DD09"
                                                     PCInc = 2
                                     Case reg_DE: codestr = "DD19"
                                                     PCInc = 2
                                     Case reg_IX: codestr = "DD29"
                                                     PCInc = 2
                                     Case reg_SP: codestr = "DD39"
                                                     PCInc = 2
                                     Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                                  End Select
                     Case reg_IY: Select Case Reg2Typ
                                     Case reg_BC: codestr = "FD09"
                                                     PCInc = 2
                                     Case reg_DE: codestr = "FD19"
                                                     PCInc = 2
                                     Case reg_IY: codestr = "FD29"
                                                     PCInc = 2
                                     Case reg_SP: codestr = "FD39"
                                                     PCInc = 2
                                     Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                                  End Select
                     Case Else: Call Error("### Bad Destination Operand - " + Reg1str)
                  End Select ' ADD
      Case "ADC": Select Case Reg1Typ
                     Case reg_A:
                        Select Case Reg2Typ
                           Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                              codestr = Hex(&H88 + Reg2Typ)
                              PCInc = 1
                           Case reg_Imed
                              codestr = "CE" + Bytes(Reg2str)
                              PCInc = 2
                           Case reg_IX_Indir:
                              codestr = "DD8E" + XYDisplacement(Reg2str)
                              PCInc = 3
                           Case reg_IY_Indir:
                              codestr = "FD8E" + XYDisplacement(Reg2str)
                              PCInc = 3
                           Case Else:  Call Error("### Bad Source Operand - " + Reg2str)
                        End Select
                     Case reg_HL:
                        PCInc = 2
                        Select Case Reg2Typ
                           Case reg_BC: codestr = "ED4A"
                           Case reg_DE: codestr = "ED5A"
                           Case reg_HL: codestr = "ED6A"
                           Case reg_SP: codestr = "ED7A"
                           Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                        End Select
                     Case Else: Call Error("### Bad Destination Operand - " + Reg1str)
                  End Select ' ADC
      Case "SUB":
                  Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = Hex(&H90 + Reg1Typ)
                        PCInc = 1
                     Case reg_Imed
                        codestr = "D6" + Bytes(Reg1str)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DD96" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case reg_IY_Indir:
                        codestr = "FD96" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case Else:  Call Error("### Bad Source Operand - " + Reg2str)
                  End Select ' SUB
       Case "SBC":
                  If Reg1Typ = reg_A Then
                     Select Case Reg2Typ
                        Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                           codestr = Hex(&H98 + Reg2Typ)
                           PCInc = 1
                        Case reg_Imed
                           codestr = "DE" + Bytes(Reg2str)
                           PCInc = 2
                        Case reg_IX_Indir:
                           codestr = "DD9E" + XYDisplacement(Reg2str)
                           PCInc = 3
                        Case reg_IY_Indir:
                           codestr = "FD9E" + XYDisplacement(Reg2str)
                           PCInc = 3
                        Case Else:  Call Error("### Bad Source Operand - " + Reg2str)
                     End Select
                  ElseIf Reg1Typ = reg_HL Then
                     PCInc = 2
                     Select Case Reg2Typ
                        Case reg_BC: codestr = "ED42"
                        Case reg_DE: codestr = "ED52"
                        Case reg_HL: codestr = "ED62"
                        Case reg_SP: codestr = "ED72"
                        Case Else:
                     End Select
                  Else
                     Call Error("### Bad Destination Operand - " + Reg1str)
                  End If  ' SBC
      Case "AND":
                  If Reg2str <> "" Then Call Error("### Extra Operand - " + Reg2str)
                  Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = Hex(&HA0 + Reg1Typ)
                        PCInc = 1
                     Case reg_Imed
                        codestr = "E6" + Bytes(Reg1str)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDA6" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case reg_IY_Indir:
                        codestr = "FDA6" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case Else
                        Call Error("### Bad Source Operand - " + Reg2str)
                  End Select ' AND
      Case "XOR":
                  If Reg2str <> "" Then Call Error("### Extra Operand - " + Reg2str)
                  Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = Hex(&HA8 + Reg1Typ)
                        PCInc = 1
                     Case reg_Imed
                        codestr = "EE" + Bytes(Reg1str)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDAE" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case reg_IY_Indir:
                        codestr = "FDAE" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case Else
                        Call Error("### Bad Source Operand - " + Reg2str)
                  End Select ' XOR
      Case "OR":
                  If Reg2str <> "" Then Call Error("### Extra Operand - " + Reg2str)
                  Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = Hex(&HB0 + Reg1Typ)
                        PCInc = 1
                     Case reg_Imed
                        codestr = "F6" + Bytes(Reg1str)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDB6" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case reg_IY_Indir:
                        codestr = "FDB6" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case Else
                        Call Error("### Bad Source Operand - " + Reg2str)
                  End Select ' OR
      Case "CP":
                  If Reg2str <> "" Then Call Error("### Extra Operand - " + Reg2str)
                  Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = Hex(&HB8 + Reg1Typ)
                        PCInc = 1
                     Case reg_Imed
                        codestr = "FE" + Bytes(Reg1str)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDBE" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case reg_IY_Indir:
                        codestr = "FDBE" + XYDisplacement(Reg1str)
                        PCInc = 3
                     Case Else
                        Call Error("### Bad Source Operand - " + Reg2str)
                  End Select ' CP
                  
      Case "POP": PCInc = 1
                  Select Case Reg1Typ
                     Case reg_BC: codestr = "C1"
                     Case reg_DE: codestr = "D1"
                     Case reg_HL: codestr = "E1"
                     Case reg_AF: codestr = "F1"
                     Case reg_IX: codestr = "DDE1"
                                  PCInc = PCInc + 1
                     Case reg_IY: codestr = "FDE1"
                                  PCInc = PCInc + 1
                     Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                  End Select ' POP
      Case "PUSH": PCInc = 1
                   Select Case Reg1Typ
                     Case reg_BC: codestr = "C5"
                     Case reg_DE: codestr = "D5"
                     Case reg_HL: codestr = "E5"
                     Case reg_AF: codestr = "F5"
                     Case reg_IX: codestr = "DDE5"
                                  PCInc = PCInc + 1
                     Case reg_IY: codestr = "FDE5"
                                  PCInc = PCInc + 1
                     Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                  End Select ' PUSH
      Case "DJNZ": PCInc = 2
                   codestr = "10" + CalcRelOffset(ASMPC + 2, Reg1str)
      Case "JR":
                  PCInc = 2
                  Select Case Reg1str
                     Case "NZ":
                        codestr = "20" + CalcRelOffset(ASMPC + 2, Reg2str)
                     Case "Z":
                        codestr = "28" + CalcRelOffset(ASMPC + 2, Reg2str)
                     Case "NC":
                        codestr = "30" + CalcRelOffset(ASMPC + 2, Reg2str)
                     Case "C":
                        codestr = "38" + CalcRelOffset(ASMPC + 2, Reg2str)
                     Case Else:
                        If Reg2Typ = reg_None Then
                           codestr = "18" + CalcRelOffset(ASMPC + 2, Reg1str)
                        Else
                           Call Error("### Invalid Conditional - " + Reg1str)
                        End If
                  End Select ' JR
      Case "RET": PCInc = 1
                  Select Case Reg1str
                     Case "NZ": codestr = "C0"
                     Case "Z": codestr = "C8"
                     Case "NC": codestr = "D0"
                     Case "C": codestr = "D8"
                     Case "PO": codestr = "E0"
                     Case "PE": codestr = "E8"
                     Case "P": codestr = "F0"
                     Case "M": codestr = "F8"
                     Case Else:
                        If Reg2str = "" Then
                           codestr = "C9"
                        Else
                           Call Error("### Bad Conditional - " + Reg1str)
                        End If
                  End Select ' RET
      Case "JP":
                 If Reg2str <> "" Then
                    If Left(Reg2str, 1) = "$" Then
                       RelOffset = ASMPC + Val(Right(Reg2str, Len(Reg2str) - 1))
                    Else
                       RelOffset = Immed16(Reg2str)
                    End If
                    PCInc = 3
                    Select Case Reg1str
                       Case "NZ": codestr = "C2" + OpcodeHex4(RelOffset)
                       Case "Z":
                                  codestr = "CA" + OpcodeHex4(RelOffset)
                       Case "NC": codestr = "D2" + OpcodeHex4(RelOffset)
                       Case "C": codestr = "DA" + OpcodeHex4(RelOffset)
                       Case "PO": codestr = "E2" + OpcodeHex4(RelOffset)
                       Case "PE": codestr = "EA" + OpcodeHex4(RelOffset)
                       Case "P": codestr = "F2" + OpcodeHex4(RelOffset)
                       Case "M": codestr = "FA" + OpcodeHex4(RelOffset)
                       Case Else: Call Error("### Bad Conditional - " + Reg1str)
                    End Select
                  ElseIf Reg2str = "" Then
                     Select Case Reg1Typ
                        Case reg_Imed:
                                   If Left(Reg1str, 1) = "$" Then
                                      RelOffset = ASMPC + Val(Right(Reg1str, Len(Reg1str) - 1))
                                      codestr = "C3" + OpcodeHex4(RelOffset)
                                   Else
                                      codestr = "C3" + OpcodeHex4(Immed16(Reg1str))
                                   End If
                                   PCInc = 3
                        Case reg_M: codestr = "E9"
                                   PCInc = 1
                        Case reg_IX_Indir: codestr = "DDE9"
                                   PCInc = 2
                        Case reg_IY_Indir: codestr = "FDE9"
                                   PCInc = 2
                        Case Else: Call Error("### Bad Operand")
                    End Select
                  Else: Call Error("### Bad Operands")
                  End If ' JP
      Case "CALL": Select Case Reg1str
                      Case "NZ": codestr = "C4" + OpcodeHex4(Immed16(Reg2str))
                      Case "Z": codestr = "CC" + OpcodeHex4(Immed16(Reg2str))
                      Case "NC": codestr = "D4" + OpcodeHex4(Immed16(Reg2str))
                      Case "C": codestr = "DC" + OpcodeHex4(Immed16(Reg2str))
                      Case "PO": codestr = "E4" + OpcodeHex4(Immed16(Reg2str))
                      Case "PE": codestr = "EC" + OpcodeHex4(Immed16(Reg2str))
                      Case "P": codestr = "F4" + OpcodeHex4(Immed16(Reg2str))
                      Case "M": codestr = "FC" + OpcodeHex4(Immed16(Reg2str))
                      Case Else:
                         If Reg2str = "" Then
                            codestr = "CD" + OpcodeHex4(Immed16(Reg1str))
                         Else
                            Call Error("### Bad Conditional - " + Reg1str)
                         End If
                   End Select ' CALL
                   PCInc = 3
      Case "RST": PCInc = 1
                  Select Case Reg1str
                     Case "0", "0H", "00H": codestr = "C7"
                     Case "8", "08H": codestr = "CF"
                     Case "16", "10H": codestr = "D7"
                     Case "24", "18H": codestr = "DF"
                     Case "32", "20H": codestr = "E7"
                     Case "40", "28H": codestr = "EF"
                     Case "48", "30H": codestr = "F7"
                     Case "56", "38H": codestr = "FF"
                     Case Else: Call Error("### Bad Operand - " + Reg1str)
                  End Select ' RST
      Case "IM":  PCInc = 2
                  Select Case Reg1str
                     Case "0": codestr = "ED46"
                     Case "1": codestr = "ED56"
                     Case "2": codestr = "ED5E"
                     Case Else: Call Error("### Bad Mode Selection - " + Reg1str)
                  End Select  ' IM
      Case "OUT": PCInc = 2
                  If (Reg1Typ = reg_Imed_Indir) And (Reg2Typ = reg_A) Then
                     codestr = "D3" + hex2L(Immed8(Reg1str))
                  ElseIf (Reg1Typ = reg_C_Indir) Then
                     Select Case Reg2Typ
                        Case reg_B: codestr = "ED41"
                        Case reg_C: codestr = "ED49"
                        Case reg_D: codestr = "ED51"
                        Case reg_E: codestr = "ED59"
                        Case reg_H: codestr = "ED61"
                        Case reg_L: codestr = "ED69"
                        Case reg_A: codestr = "ED79"
                        Case Else: Call Error("### Bad Source Operand - " + Reg2str)
                     End Select
                  Else: Call Error("### Bad Destination Operand - " + Reg1str)
                  End If ' OUT
      Case "IN":  PCInc = 2
                  If (Reg1Typ = reg_A) And (Reg2Typ = reg_Imed_Indir) Then
                     codestr = "DB" + hex2L(Immed8(Reg2str))
                  ElseIf (Reg2Typ = reg_C_Indir) Then
                     codestr = "ED" + hex2L(&H40 + Reg1Typ * 8)
                     Else: Call Error("### Bad Destination Operand - " + Reg1str)
                  End If ' IN
      Case "RLC": Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = "CB" + hex2L(Reg1Typ)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDCB" + XYDisplacement(Reg1str) + "06"
                        PCInc = 4
                     Case reg_IY_Indir:
                        codestr = "FDCB" + XYDisplacement(Reg1str) + "06"
                        PCInc = 4
                     Case Else
                        Call Error("### Bad Operand - " + Reg1str)
                  End Select ' RLC
      Case "RRC": Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = "CB" + hex2L(Reg1Typ + 8)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDCB" + XYDisplacement(Reg1str) + "0E"
                        PCInc = 4
                     Case reg_IY_Indir:
                        codestr = "FDCB" + XYDisplacement(Reg1str) + "0E"
                        PCInc = 4
                     Case Else
                        Call Error("### Bad Operand - " + Reg1str)
                  End Select ' RRC
      Case "RL": Select Case Reg1Typ
                    Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                       codestr = "CB" + hex2L(Reg1Typ + &H10)
                       PCInc = 2
                    Case reg_IX_Indir:
                       codestr = "DDCB" + XYDisplacement(Reg1str) + "16"
                       PCInc = 4
                    Case reg_IY_Indir:
                       codestr = "FDCB" + XYDisplacement(Reg1str) + "16"
                       PCInc = 4
                    Case Else
                       Call Error("### Bad Operand - " + Reg1str)
                 End Select ' RL
      Case "RR": Select Case Reg1Typ
                    Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                       codestr = "CB" + hex2L(Reg1Typ + &H18)
                       PCInc = 2
                    Case reg_IX_Indir:
                       codestr = "DDCB" + XYDisplacement(Reg1str) + "1E"
                       PCInc = 4
                    Case reg_IY_Indir:
                       codestr = "FDCB" + XYDisplacement(Reg1str) + "1E"
                       PCInc = 4
                    Case Else
                       Call Error("### Bad Operand - " + Reg1str)
                 End Select ' RR
      Case "SLA": Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = "CB" + hex2L(Reg1Typ + &H20)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDCB" + XYDisplacement(Reg1str) + "26"
                        PCInc = 4
                     Case reg_IY_Indir:
                        codestr = "FDCB" + XYDisplacement(Reg1str) + "26"
                        PCInc = 4
                     Case Else
                        Call Error("### Bad Operand - " + Reg1str)
                  End Select ' SLA
      Case "SRA": Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = "CB" + hex2L(Reg1Typ + &H28)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDCB" + XYDisplacement(Reg1str) + "2E"
                        PCInc = 4
                     Case reg_IY_Indir:
                        codestr = "FDCB" + XYDisplacement(Reg1str) + "2E"
                        PCInc = 4
                     Case Else
                        Call Error("### Bad Operand - " + Reg1str)
                  End Select ' SRA
      Case "SRL": Select Case Reg1Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = "CB" + hex2L(Reg1Typ + &H38)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDCB" + XYDisplacement(Reg1str) + "3E"
                        PCInc = 4
                     Case reg_IY_Indir:
                        codestr = "FDCB" + XYDisplacement(Reg1str) + "3E"
                        PCInc = 4
                     Case Else
                        Call Error("### Bad Operand - " + Reg1str)
                  End Select ' SRL
      Case "BIT":
                  If (Reg1str >= "0") And (Reg1str <= "9") Then
                  Select Case Reg2Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = "CB" + hex2L(&H40 + Val(Reg1str) * 8 + Reg2Typ)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDCB" + XYDisplacement(Reg2str) + hex2L(&H40 + Val(Reg1str) * 8 + 6)
                        PCInc = 4
                     Case reg_IY_Indir:
                        codestr = "FDCB" + XYDisplacement(Reg2str) + hex2L(&H40 + Val(Reg1str) * 8 + 6)
                        PCInc = 4
                     Case Else
                        Call Error("### Bad Register - " + Reg2str)
                     End Select
                  Else
                     Call Error("### Bad Bit Number (0-8) - " + Reg1str)
                  End If ' BIT
      Case "RES":
                  If (Reg1str >= "0") And (Reg1str <= "9") Then
                  Select Case Reg2Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = "CB" + hex2L(&H80 + Val(Reg1str) * 8 + Reg2Typ)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDCB" + XYDisplacement(Reg2str) + hex2L(&H80 + Val(Reg1str) * 8 + 6)
                        PCInc = 4
                     Case reg_IY_Indir:
                        codestr = "FDCB" + XYDisplacement(Reg2str) + hex2L(&H80 + Val(Reg1str) * 8 + 6)
                        PCInc = 4
                     Case Else
                        Call Error("### Bad Register")
                     End Select
                  Else
                     Call Error("### Bad Bit Number (0-8) - " + Reg1str)
                  End If ' RES
      Case "SET":
                  If (Reg1str >= "0") And (Reg1str <= "9") Then
                  Select Case Reg2Typ
                     Case reg_B, reg_C, reg_D, reg_E, reg_H, reg_L, reg_M, reg_A:
                        codestr = "CB" + hex2L(&HC0 + Val(Reg1str) * 8 + Reg2Typ)
                        PCInc = 2
                     Case reg_IX_Indir:
                        codestr = "DDCB" + XYDisplacement(Reg2str) + hex2L(&HC0 + Val(Reg1str) * 8 + 6)
                        PCInc = 4
                     Case reg_IY_Indir:
                        codestr = "FDCB" + XYDisplacement(Reg2str) + hex2L(&HC0 + Val(Reg1str) * 8 + 6)
                        PCInc = 4
                     Case Else
                        Call Error("### Bad Register - " + Reg2str)
                     End Select
                  Else
                     Call Error("### Bad Bit Number (0-8) - " + Reg1str)
                  End If ' SET
      Case Else: Call Error("### Undefined Opcode - " + OpcodeStr)
   End Select
   Assemble = Hex4(ASMPC) + codestr
End Function

