Attribute VB_Name = "modDissasm"
Option Explicit


Private Function exeRead(PC As Long) As Long
   exeRead = RAM(PC)
   PC = (PC + 1)
   PC = PC And CLng("&HFFFF")
End Function

Private Function AddLabel(address As Long) As String
Dim workLong As Long
   
   address = address And CLng("&HFFFF")
   If Memory(address).Label = "" Then
      Memory(address).Label = "L" + Right("0000" + Hex(address), 4)
      AddLabel = "L" + Right("0000" + Hex(address), 4)
   Else
      AddLabel = Memory(address).Label
   End If
End Function

' Dissassembles the next opcode.
' PC contains the address for the next opcode.
' PC points to the next opcode address when done.
' will read as many bytes as needed for the given opcode
' and update PC to point to the next opcode on exit.
' Address and machine codes are returned in MachCode
' label and source code is passed back in SrcCode.
Public Sub DissAssemble(PC As Long, MachCode As String, SrcCode As String)
Dim OpCode As Integer
Dim data1 As Long
Dim Data2 As Long
Dim Data3 As Long
Dim Label As String
   
   If PC > EndAddress Then
      MachCode = ""
      SrcCode = ""
      PC = PC + 1
      Exit Sub
   End If
   MachCode = Right("0000" + Hex(PC), 4) + " "
   Label = Left(Memory(PC).Label + "                ", 8)
   

   Select Case Memory(PC).Usage
      Case 1: OperMode = 1            ' ASCII
      Case 2: OperMode = 2            ' BYTE
      Case 3: OperMode = 3            ' INSTRUCTION
      Case 4: OperMode = 4            ' STORAGE
      Case 5: OperMode = 5            ' WORD
      Case 6: OperMode = 6            ' END
   End Select
   
   Select Case OperMode
      Case 1: Call DisplayASCII(PC, MachCode, SrcCode)
      Case 2: Call DisplayByte(PC, MachCode, SrcCode)
      Case 3: OpCode = exeRead(PC)
              MachCode = MachCode + Right("00" + Hex(OpCode), 2)
              Select Case OpCode
                 Case &H0: SrcCode = "NOP"
                 Case &H1: data1 = exeRead(PC)
                           Data2 = exeRead(PC)
                           MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                           SrcCode = "LD    BC," + Right("00" + Hex(Data2), 2) + Right("00" + Hex(data1), 2) + "H"
                           AddLabel (CLng(Data2 * CLng(256) + data1))
                 Case &H2: SrcCode = "LD    (BC),A"
                 Case &H3: SrcCode = "INC   BC"
                 Case &H4: SrcCode = "INC   B"
                 Case &H5: SrcCode = "DEC   B"
                 Case &H6: data1 = exeRead(PC)
                           MachCode = MachCode + Right("00" + Hex(data1), 2)
                           SrcCode = "LD    B," + Right("00" + Hex(data1), 2) + "H"
                 Case &H7: SrcCode = "RLCA"
                 Case &H8: SrcCode = "EX    AF,AF'"
                 Case &H9: SrcCode = "ADD   HL,BC"
                 Case &HA: SrcCode = "LD    A,(BC)"
                 Case &HB: SrcCode = "DEC   BC"
                 Case &HC: SrcCode = "INC   C"
                 Case &HD: SrcCode = "DEC   C"
                 Case &HE: data1 = exeRead(PC)
                           MachCode = MachCode + Right("00" + Hex(data1), 2)
                           SrcCode = "LD    C," + Right("00" + Hex(data1), 2) + "H"
                 Case &HF: SrcCode = "RRCA"
                 Case &H10: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "DJNZ  "
                            If data1 > 127 Then data1 = data1 - 256
                            SrcCode = SrcCode + AddLabel(CLng(PC + data1)) 'Hex4(PC + data1) + "H"
                 Case &H11: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "LD    DE," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &H12: SrcCode = "LD    (DE),A"
                 Case &H13: SrcCode = "INC   DE"
                 Case &H14: SrcCode = "INC   D"
                 Case &H15: SrcCode = "DEC   D"
                 Case &H16: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "LD    D," + Right("00" + Hex(data1), 2)
                 Case &H17: SrcCode = "RLA"
                 Case &H18: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "JR    "
                            If data1 > 127 Then data1 = data1 - 256
                            SrcCode = SrcCode + AddLabel(CLng(PC + data1))
                 Case &H19: SrcCode = "ADD   HL,DE"
                 Case &H1A: SrcCode = "LD    A,(DE)"
                 Case &H1B: SrcCode = "DEC   DE"
                 Case &H1C: SrcCode = "INC   E"
                 Case &H1D: SrcCode = "DEC   E"
                 Case &H1E: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "LD    E," + Right("00" + Hex(data1), 2)
                 Case &H1F: SrcCode = "RRA"
                 Case &H20: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "JR    NZ,"
                            If data1 > 127 Then data1 = data1 - 256
                            SrcCode = SrcCode + AddLabel(CLng(PC + data1))
                 Case &H21: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "LD    HL," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &H22: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "LD    (" + AddLabel(CLng(Data2 * CLng(256) + data1)) + "),HL"
                 Case &H23: SrcCode = "INC   HL"
                 Case &H24: SrcCode = "INC   H"
                 Case &H25: SrcCode = "DEC   H"
                 Case &H26: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "LD    H," + Right("00" + Hex(data1), 2)
                 Case &H27: SrcCode = "DAA"
                 Case &H28: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            If data1 > 127 Then data1 = data1 - 256
                            SrcCode = "JR    Z," + AddLabel(CLng(PC + data1))
                 Case &H29: SrcCode = "ADD   HL,HL"
                 Case &H2A: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "LD    HL,(" + AddLabel(CLng(Data2 * CLng(256) + data1)) + ")"
                 Case &H2B: SrcCode = "DEC   HL"
                 Case &H2C: SrcCode = "INC   L"
                 Case &H2D: SrcCode = "DEC   L"
                 Case &H2E: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "LD    L," + Right("00" + Hex(data1), 2)
                 Case &H2F: SrcCode = "CPL"
                 Case &H30: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            If data1 > 127 Then data1 = data1 - 256
                            SrcCode = "JR    NC," + AddLabel(CLng(PC + data1))
                 Case &H31: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "LD    SP," + Right("00" + Hex(Data2), 2) + Right("00" + Hex(data1), 2) + "H"
                 Case &H32: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "LD    (" + AddLabel(CLng(Data2 * CLng(256) + data1)) + "),A"
                 Case &H33: SrcCode = "INC   SP"
                 Case &H34: SrcCode = "INC   (HL)"
                 Case &H35: SrcCode = "DEC   (HL)"
                 Case &H36: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "LD    (HL)," + Right("00" + Hex(data1), 2)
                 Case &H37: SrcCode = "SCF"
                 Case &H38: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(Data2), 2) + Right("00" + Hex(data1), 2)
                            If data1 > 127 Then data1 = data1 - 256
                            SrcCode = "JR    C," + AddLabel(CLng(PC + data1))
                            
                 Case &H39: SrcCode = "ADD   HL,SP"
                 Case &H3A: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "LD    A,(" + AddLabel(CLng(Data2 * CLng(256) + data1)) + ")"
                 Case &H3B: SrcCode = "DEC   SP"
                 Case &H3C: SrcCode = "INC   A"
                 Case &H3D: SrcCode = "DEC   A"
                 Case &H3E: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "LD    A," + Right("00" + Hex(data1), 2) + "H"
                 Case &H3F: SrcCode = "CCF"
                 Case &H40, &H41, &H42, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &H4A, &H4B, &H4C, &H4D, &H4E, &H4F, _
                      &H50, &H51, &H52, &H53, &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &H5B, &H5C, &H5D, &H5E, &H5F, _
                      &H60, &H61, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H6A, &H6B, &H6C, &H6D, &H6E, &H6F, _
                      &H70, &H71, &H72, &H73, &H74, &H75, &H77, &H78, &H79, &H7A, &H7B, &H7C, &H7D, &H7E, &H7F:
                            data1 = (OpCode And &H38) / 8
                            Data2 = OpCode And &H7
                            SrcCode = "LD    " + Trim(Mid(RegList, data1 * 4 + 1, 4)) + "," + Trim(Mid(RegList, Data2 * 4 + 1, 4))
                 Case &H76: SrcCode = "HALT  "
                 Case &H80, &H81, &H82, &H83, &H84, &H85, &H86, &H87:
                            data1 = (OpCode And &H7)
                            SrcCode = "ADD   A," + Trim(Mid(RegList, data1 * 4 + 1, 4))
                 Case &H88, &H89, &H8A, &H8B, &H8C, &H8D, &H8E, &H8F:
                            data1 = (OpCode And &H7)
                            SrcCode = "ADC   A," + Trim(Mid(RegList, data1 * 4 + 1, 4))
                 Case &H90, &H91, &H92, &H93, &H94, &H95, &H96, &H97:
                            data1 = (OpCode And &H7)
                            SrcCode = "SUB   " + Trim(Mid(RegList, data1 * 4 + 1, 4))
                 Case &H98, &H99, &H9A, &H9B, &H9C, &H9D, &H9E, &H9F:
                            data1 = (OpCode And &H7)
                            SrcCode = "SBC   A," + Trim(Mid(RegList, data1 * 4 + 1, 4))
                 Case &HA0, &HA1, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7:
                            data1 = (OpCode And &H7)
                            SrcCode = "AND   " + Trim(Mid(RegList, data1 * 4 + 1, 4))
                 Case &HA8, &HA9, &HAA, &HAB, &HAC, &HAD, &HAE, &HAF:
                            data1 = (OpCode And &H7)
                            SrcCode = "XOR   " + Trim(Mid(RegList, data1 * 4 + 1, 4))
                 Case &HB0, &HB1, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7:
                            data1 = (OpCode And &H7)
                            SrcCode = "OR    " + Trim(Mid(RegList, data1 * 4 + 1, 4))
                 Case &HB8, &HB9, &HBA, &HBB, &HBC, &HBD, &HBE, &HBF:
                            data1 = (OpCode And &H7)
                            SrcCode = "CP    " + Trim(Mid(RegList, data1 * 4 + 1, 4))
                 Case &HC0: SrcCode = "RET   NZ"
                 Case &HC1: SrcCode = "POP   BC"
                 Case &HC2: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    NZ," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HC3: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    " + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HC4: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  NZ," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HC5: SrcCode = "PUSH  BC"
                 Case &HC6: data1 = exeRead(PC)
                            SrcCode = "ADD    A," + Right("00" + Hex(data1), 2)
                 Case &HC7: SrcCode = "RST   00H"
                 Case &HC8: SrcCode = "RET   Z"
                 Case &HC9: SrcCode = "RET"
                 Case &HCA: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    Z," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HCB: Call CBopcodes(PC, MachCode, SrcCode)
                 Case &HCC: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  Z," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HCD: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  " + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HCE: data1 = exeRead(PC)
                            SrcCode = "ADC   A," + Right("00" + Hex(data1), 2)
                 Case &HCF: SrcCode = "RST   08H"
                 Case &HD0: SrcCode = "RET   NC"
                 Case &HD1: SrcCode = "POP   DE"
                 Case &HD2: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    NC," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HD3: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "OUT   (" + Right("00" + Hex(data1), 2) + "H),A"
                 Case &HD4: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  NC," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HD5: SrcCode = "PUSH  DE"
                 Case &HD6: data1 = exeRead(PC)
                            SrcCode = "SUB   " + Right("00" + Hex(data1), 2)
                 Case &HD7: SrcCode = "RST   10H   ;16"
                 Case &HD8: SrcCode = "RET   C"
                 Case &HD9: SrcCode = "EXX"
                 Case &HDA: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    C," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HDB: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "OUT   A,(" + Right("00" + Hex(data1), 2) + "H)"
                 Case &HDC: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  C," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HDD: Call DDFDOpcodes(PC, MachCode, SrcCode, "IX")
                 Case &HDE: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "SBC   A," + Right("00" + Hex(data1), 2)
                 Case &HDF: SrcCode = "RST   18H   ;24"
                 Case &HE0: SrcCode = "RET   PO"
                 Case &HE1: SrcCode = "POP   HL"
                 Case &HE2: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    PO," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HE3: SrcCode = "EX    (SP),HL"
                 Case &HE4: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  PO," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HE5: SrcCode = "PUSH  HL"
                 Case &HE6: data1 = exeRead(PC)
                            SrcCode = "AND   " + Right("00" + Hex(data1), 2)
                 Case &HE7: SrcCode = "RST   20H   ;32"
                 Case &HE8: SrcCode = "RET   PE"
                 Case &HE9: SrcCode = "JP    (HL)"
                 Case &HEA: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    PE," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HEB: SrcCode = "EX    DE,HL"
                 Case &HEC: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  PE," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HED: Call EDOpcodes(PC, MachCode, SrcCode)
                 Case &HEE: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "XOR   " + Right("00" + Hex(data1), 2)
                 Case &HEF: SrcCode = "RST   28H   ;40"
                 Case &HF0: SrcCode = "RET   P"
                 Case &HF1: SrcCode = "POP   AF"
                 Case &HF2: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    P," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HF3: SrcCode = "DI"
                 Case &HF4: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  P," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HF5: SrcCode = "PUSH  AF"
                 Case &HF6: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "OR    " + Right("00" + Hex(data1), 2)
                 Case &HF7: SrcCode = "RST   30H   ;48"
                 Case &HF8: SrcCode = "RET   M"
                 Case &HF9: SrcCode = "LD    SP,HL"
                 Case &HFA: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "JP    M," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HFB: SrcCode = "EI"
                 Case &HFC: data1 = exeRead(PC)
                            Data2 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                            SrcCode = "CALL  M," + AddLabel(CLng(Data2 * CLng(256) + data1))
                 Case &HFD: Call DDFDOpcodes(PC, MachCode, SrcCode, "IY")
                 Case &HFE: data1 = exeRead(PC)
                            MachCode = MachCode + Right("00" + Hex(data1), 2)
                            SrcCode = "CP    " + Right("00" + Hex(data1), 2)
                 Case &HFF: SrcCode = "RST   38H   ;56"
                 Case Else: MachCode = MachCode + " !BAD OPCODE"
              End Select
      Case 4: Call DisplayStorage(PC, MachCode, SrcCode)
      Case 5: Call DisplayWord(PC, MachCode, SrcCode)
      Case 6: EndAddress = PC
              SrcCode = "        END"
              PC = PC + 1
   End Select
   SrcCode = Left(Label + "          ", 8) + SrcCode

End Sub

Private Sub CBopcodes(PC As Long, MachCode As String, SrcCode As String)
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim oper As String
   
   OpCode = exeRead(PC)
   MachCode = MachCode + Right("00" + Hex(OpCode), 2)
   data1 = (OpCode And &HF0) / 16
   
   Select Case data1
      Case &H0, &H1, &H2, &H3:
                Data3 = (OpCode And &H38) / 8
                Select Case Data3
                   Case 0: SrcCode = "RLC   "
                   Case 1: SrcCode = "RRC   "
                   Case 2: SrcCode = "RL    "
                   Case 3: SrcCode = "RR    "
                   Case 4: SrcCode = "SLA   "
                   Case 5: SrcCode = "SRA   "
                   Case 6: SrcCode = "illop "
                   Case 7: SrcCode = "SRL   "
                End Select
                Data2 = OpCode And &H7
                SrcCode = SrcCode + Trim(Mid(RegList, Data2 * 4 + 1, 4))
      Case &H4, &H5, &H6, &H7:
                data1 = (OpCode And &H38) / 8
                Data2 = OpCode And &H7
                SrcCode = "BIT   " + Right("00" + data1, 1) + "," + Trim(Mid(RegList, Data2 * 4 + 1, 4))
      Case &H8, &H9, &HA, &HB:
                data1 = (OpCode And &H38) / 8
                Data2 = OpCode And &H7
                SrcCode = "RES   " + Right("00" + data1, 1) + "," + Trim(Mid(RegList, Data2 * 4 + 1, 4))
      Case &HC, &HD, &HE, &HF:
                data1 = (OpCode And &H38) / 8
                Data2 = OpCode And &H7
                SrcCode = "SET   " + Right("00" + data1, 1) + "," + Trim(Mid(RegList, Data2 * 4 + 1, 4))
      Case Else: SrcCode = " @BAD OPCODE"
      End Select
End Sub

Private Sub DDFDCBOpcodes(PC As Long, MachCode As String, SrcCode As String, IXIY As String)
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim dat2 As Integer
Dim RegID As Boolean
   
   If IXIY = "IX" Then RegID = True Else RegID = False
   data1 = exeRead(PC)
   OpCode = exeRead(PC)
   MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(OpCode), 2) + "  "
   Select Case OpCode
      Case &H6: SrcCode = "RLC   "
      Case &HE: SrcCode = "RRC   "
      Case &H16: SrcCode = "RL    "
      Case &H1E: SrcCode = "RR    "
      Case &H26: SrcCode = "SLA   "
      Case &H2E: SrcCode = "SRA   "
      Case &H46, &H4E, &H56, &H5E, &H66, &H6E, &H76, &H7E:
                 Data2 = (OpCode And &H38) \ 8
                 SrcCode = "BIT   " + Hex(Data2) + ","
      Case &H86, &H8E, &H96, &H9E, &HA6, &HAE, &HB6, &HBE:
                 Data2 = (OpCode And &H38) \ 8
                 SrcCode = "RES   " + Hex(Data2) + ","
      Case &HC6, &HCE, &HD6, &HDE, &HE6, &HEE, &HF6, &HFE:
                 Data2 = (OpCode And &H38) \ 8
                 SrcCode = "SET   " + Hex(Data2) + ","
      Case Else: SrcCode = " #BAD OPCODE  "
   End Select
   SrcCode = SrcCode + "(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
End Sub

Private Sub DDFDOpcodes(PC As Long, MachCode As String, SrcCode As String, IXIY As String)
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim dat2 As Integer
Dim RegID As Integer

   OpCode = exeRead(PC)
   MachCode = MachCode + Right("00" + Hex(OpCode), 2)
   Select Case OpCode
      Case &H9: SrcCode = "ADD   " + IXIY + ",BC"
      Case &H19: SrcCode = "ADD   " + IXIY + ",DE"
      Case &H21: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    " + IXIY + "," + AddLabel(CLng(Data2 * CLng(256) + data1))
      Case &H22: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    (" + AddLabel(CLng(Data2 * CLng(256) + data1)) + ")," + IXIY
      Case &H23: SrcCode = "INC   " + IXIY
      Case &H29: SrcCode = "ADD   " + IXIY + "," + IXIY
      Case &H2A: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    " + IXIY + ",(" + AddLabel(CLng(Data2 * CLng(256) + data1)) + ")"
      Case &H2B: SrcCode = "DEC   " + IXIY
      Case &H34: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "INC   (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H35: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "DEC   (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H36: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)," + Right("00" + Hex(Data2), 2)
      Case &H39: SrcCode = "ADD   " + IXIY + ",SP"
      Case &H46: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    B,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H4E: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    C,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H56: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    D,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H5E: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    E,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H66: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    H,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H6E: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    L,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H70: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H),B"
      Case &H71: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H),C"
      Case &H72: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H),D"
      Case &H73: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H),E"
      Case &H74: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H),H"
      Case &H75: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H),L"
      Case &H77: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H),A"
      Case &H7E: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "LD    A,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H86: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "ADD   A,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H8E: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "ADC   A,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H96: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "SUB   (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &H9E: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "SBC   A,(" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &HA6: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "AND   (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &HAE: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "XOR   (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &HB6: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "OR    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &HBE: data1 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2)
                 SrcCode = "CP    (" + IXIY + "+" + Right("00" + Hex(data1), 2) + "H)"
      Case &HCB: Call DDFDCBOpcodes(PC, MachCode, SrcCode, IXIY)
      Case &HE1: SrcCode = "POP   " + IXIY
      Case &HE3: SrcCode = "EX    (SP)," + IXIY
      Case &HE5: SrcCode = "PUSH  " + IXIY
      Case &HE9: SrcCode = "JP    " + IXIY
      Case &HF9: SrcCode = "LD    SP," + IXIY
      Case Else: SrcCode = " $BAD OPCODE"
   End Select
End Sub

Private Sub EDOpcodes(PC As Long, MachCode As String, SrcCode As String)
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim dat2 As Integer
   OpCode = exeRead(PC)
   MachCode = MachCode + Right("00" + Hex(OpCode), 2)
   Select Case OpCode
      Case &H40: SrcCode = "IN    B,(C)"
      Case &H41: SrcCode = "OUT   (C),B"
      Case &H42: SrcCode = "SBC   HL,BC"
      Case &H43: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    (" + AddLabel(CLng(Data2 * CLng(256) + data1)) + "),BC"
      Case &H44: SrcCode = "NEG"
      Case &H45: SrcCode = "RETN"
      Case &H46: SrcCode = "IM    0"
      Case &H47: SrcCode = "LD    I,A"
      Case &H48: SrcCode = "IN    C,(C)"
      Case &H49: SrcCode = "OUT   (C),C"
      Case &H4A: SrcCode = "ADC   HL,BC"
      Case &H4B: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    BC,(" + AddLabel(CLng(Data2 * CLng(256) + data1)) + ")"
      Case &H4D: SrcCode = "RETI"
      Case &H4F: SrcCode = "LD    R,A"
      Case &H50: SrcCode = "IN    D,(C)"
      Case &H51: SrcCode = "OUT   (C),D"
      Case &H52: SrcCode = "SBC   HL,DE"
      Case &H53: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    (" + AddLabel(CLng(Data2 * CLng(256) + data1)) + "),DE"
      Case &H56: SrcCode = "IM    1"
      Case &H57: SrcCode = "LD    A,I"
      Case &H58: SrcCode = "IN    E,(C)"
      Case &H59: SrcCode = "OUT   (C),E"
      Case &H5A: SrcCode = "ADC   HL,DE"
      Case &H5B: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    DE,(" + AddLabel(CLng(Data2 * CLng(256) + data1)) + ")"
      Case &H5E: SrcCode = "IM2"
      Case &H5F: SrcCode = "LD    R,A"
      Case &H60: SrcCode = "IN    H,(C)"
      Case &H61: SrcCode = "OUT   (C),H"
      Case &H62: SrcCode = "SBC   HL,HL"
      Case &H67: SrcCode = "RRD"
      Case &H68: SrcCode = "IN    L,(C)"
      Case &H69: SrcCode = "OUT   L,(C)"
      Case &H6A: SrcCode = "ADC   HL,HL"
      Case &H6F: SrcCode = "RLD"
      Case &H72: SrcCode = "SBC   HL,SP"
      Case &H73: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    (" + AddLabel(CLng(Data2 * CLng(256) + data1)) + "),SP"
      Case &H78: SrcCode = "IN    A,(C)"
      Case &H79: SrcCode = "OUT   (C),A"
      Case &H7A: SrcCode = "ADC   HL,SP"
      Case &H7B: data1 = exeRead(PC)
                 Data2 = exeRead(PC)
                 MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
                 SrcCode = "LD    SP,(" + AddLabel(CLng(Data2 * CLng(256) + data1)) + ")"
      Case &HA0: SrcCode = "LDI"
      Case &HA1: SrcCode = "CPI"
      Case &HA2: SrcCode = "INI"
      Case &HA3: SrcCode = "OUTI"
      Case &HA8: SrcCode = "LDD"
      Case &HA9: SrcCode = "CPD"
      Case &HAA: SrcCode = "IND"
      Case &HAB: SrcCode = "OUTD"
      Case &HB0: SrcCode = "LDIR"
      Case &HB1: SrcCode = "CPIR"
      Case &HB2: SrcCode = "INIR"
      Case &HB3: SrcCode = "OTIR"
      Case &HB8: SrcCode = "LDDR"
      Case &HB9: SrcCode = "CPDR"
      Case &HBA: SrcCode = "INDR"
      Case &HBB: SrcCode = "OTDR"
      Case Else: SrcCode = " %BAD OPCODE"
   End Select

End Sub

Public Sub DisplayASCII(PC As Long, MachCode As String, SrcCode As String)
Dim CharCt As Integer
Dim data1 As Byte
Dim Label As String
Dim CommaFlg As Boolean

   CommaFlg = False
   CharCt = 1
   Label = Memory(PC).Label
   MachCode = Right("0000" + Hex(PC), 4)
   data1 = exeRead(PC)
   SrcCode = "DEFB  "
   If (data1 > 31) And (data1 < 127) Then
      SrcCode = SrcCode + "'" + Chr(data1) + "'"
   Else
      SrcCode = SrcCode + Right("00" + Hex(data1), 2) + "H"
   End If
   While (CharCt < 6) And (Memory(PC).Label = "") And (Memory(PC).Usage = 0)
      data1 = exeRead(PC)
      If (data1 > 31) And (data1 < 127) Then
         SrcCode = SrcCode + ", '" + Chr(data1) + "'"
      Else
         SrcCode = SrcCode + ", " + Right("00" + Hex(data1), 2) + "H"
      End If
      CharCt = CharCt + 1
   Wend
      
      
End Sub

Public Sub DisplayByte(PC As Long, MachCode As String, SrcCode As String)
Dim data1 As Byte
Dim Label As String

   MachCode = Right("0000" + Hex(PC), 4) + " "
   Label = Memory(PC).Label
   data1 = exeRead(PC)
   MachCode = MachCode + Right("00" + Hex(data1), 2)
   SrcCode = "DEFB  " + Right("00" + Hex(data1), 2) + "H"
End Sub

Public Sub DisplayStorage(PC As Long, MachCode As String, SrcCode As String)
Dim data1 As Byte
Dim ByteCt As Long
   data1 = exeRead(PC)
   ByteCt = 1
   While (Memory(PC).Label = "") And (Memory(PC).Usage = 0) And (PC <= CLng("&H10000"))
      data1 = exeRead(PC)
      ByteCt = ByteCt + 1
   Wend
   SrcCode = "DEFS  " + Hex(ByteCt) + "H"
      
End Sub

Public Sub DisplayWord(PC As Long, MachCode, SrcCode As String)
Dim data1 As Byte
Dim Data2 As Byte
Dim Label As String

   MachCode = Right("0000" + Hex(PC), 4) + " "
   Label = Memory(PC).Label
   data1 = exeRead(PC)
   Data2 = exeRead(PC)
   MachCode = MachCode + Right("00" + Hex(data1), 2) + Right("00" + Hex(Data2), 2)
   SrcCode = "DEFW  " + Right("00" + Hex(Data2), 2) + Right("00" + Hex(data1), 2) + "H"
End Sub
