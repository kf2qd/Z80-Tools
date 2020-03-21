Attribute VB_Name = "modZ80emu"
Option Explicit

Private Type MemType
   
   Usage As Byte
   Label As String
End Type

Public Memory(65535) As MemType
Public RAM(65535) As Byte
Public OperMode As Byte
Public Updatelabels As Boolean

Private Type REG8Type
   A As Integer
   F As Integer
   B As Integer
   C As Integer
   D As Integer
   E As Integer
   H As Integer
   L As Integer
   S As Integer
   P As Integer
   I As Integer
   R As Integer
   M As Integer
End Type

Private Type FlagsType
   C As Boolean
   Z As Boolean
   P As Boolean
   S As Boolean
   N As Boolean
   H As Boolean
End Type

Private Type Reg16Type
   IX As Long
   IY As Long
   SP As Long
   PC As Long
End Type

Public Regs8 As REG8Type
Public AltRegs8 As REG8Type
Public Regs16 As Reg16Type
Public Flags As FlagsType

Public Const RegList = "B   C   D   E   H   L   (HL)A     "
Public Halted As Boolean

Public EndAddress As Long


Private Function NextByte()
   NextByte = RAM(Regs16.PC)
   Regs16.PC = (Regs16.PC + 1)
   Regs16.PC = Regs16.PC And &H10FFFF
End Function

Private Sub Write8(high As Integer, low As Integer, value As Integer)
   RAM(CLng(high) * 256 + low) = value
End Sub

Private Sub Write16(ValLo As Integer, valHi As Integer, RegHi As Integer, RegLo As Integer)
   RAM((CLng(valHi) * 256) + ValLo) = RegLo
   RAM(((CLng(valHi) * 256) + ValLo + 1) And CLng("&HFFFF")) = RegHi
End Sub

Private Function Read8(high As Integer, low As Integer) As Integer
      Read8 = RAM(CLng(high) * 256 + low)
End Function

Private Function Read16(ValLo As Integer, valHi As Integer) As Long
   Read16 = RAM((CLng(valHi) * 256) + ValLo) + CLng(RAM(((CLng(valHi) * 256) + ValLo + 1) And CLng("&HFFFF"))) * 256
End Function

Private Sub Push(val As Long)
   Regs16.SP = (Regs16.SP - 1) And CLng("&HFFFF")
   RAM(Regs16.SP) = (val And CLng("&HFF00")) / 256
   Regs16.SP = (Regs16.SP - 1) And CLng("&HFFFF")
   RAM(Regs16.SP) = val And &HFF
End Sub

Private Function pop() As Long
Dim hiBits As Long
   hiBits = RAM((Regs16.SP + 1) And CLng("&HFFFF")) * CLng(256)
   pop = RAM(Regs16.SP) + hiBits
   Regs16.SP = (Regs16.SP + 2) And CLng("&HFFFF")
End Function

Private Function Incr8(reg As Integer) As Integer
Dim test As Integer
   
   Flags.N = False                      ' an 8-bit add
   test = (reg And &HFF) + 1
   Flags.H = ((reg And &HF) = 15)
   Flags.P = (reg + 1) = 128
   test = reg + 1
   Flags.C = False
   test = test And &HFF
   Flags.S = (test And &H80) > 0
   Flags.Z = (test = 0)
   Incr8 = test
End Function

Private Sub Incr16(high As Integer, low As Integer)
Dim tempReg As Long
   
   tempReg = high
   tempReg = tempReg * 256 + low
   tempReg = (tempReg + 1) And CLng(&HFFFF)
   high = (tempReg And CLng(&HFF00)) / &HFF
   low = tempReg And &HFF
End Sub

Private Function Decr8(ByRef reg As Integer) As Integer

   Flags.N = True                    ' an 8-bit subtract
   Flags.H = (reg And &HF) = 0
   Flags.P = reg = 128
   Flags.C = False
   reg = (reg - 1) And &HFF
   Flags.S = (reg And &H80) > 0
   Flags.Z = (reg = 0)
   Decr8 = reg
End Function

Private Function Decr16(high As Integer, low As Integer)
Dim tempReg As Long
      
   tempReg = high
   tempReg = tempReg * 256 + low
   tempReg = (tempReg - 1) And CLng(&HFFFF)
      high = (tempReg And CLng("&HFF00")) / 256
   low = tempReg And &HFF
End Function

Private Sub Add8(reg As Integer)
Dim temp As Integer
   
   Flags.N = False                    ' an 8 bit add
   Flags.H = (((Regs8.A And &HF) + (reg And &HF)) And &H10) > 0
   Flags.P = (((Regs8.A And &H7F) + (reg And &H7F)) And &H80) > 0
   Regs8.A = Regs8.A + reg
   Flags.S = (Regs8.A And &H80) > 0
   Flags.C = (Regs8.A And &H100) > 0
   Flags.P = Flags.P Xor Flags.C
   Flags.Z = (Regs8.A And &HFF) = 0
   Regs8.A = Regs8.A And &HFF
End Sub

Private Sub Adc8(reg As Integer)
Dim temp As Integer
   
   
   Flags.N = False                    ' an 8 bit add
   Flags.H = (((Regs8.A And &HF) + (reg And &HF)) And &H10) > 0
   Flags.P = (((Regs8.A And &H7F) + (reg And &H7F)) And &H80) > 0
   Regs8.A = Regs8.A + reg
   If Flags.C Then Regs8.A = Regs8.A + 1
   Flags.S = (Regs8.A And &H80) > 0
   Flags.C = (Regs8.A And &H100) > 0
   Flags.P = Flags.P Xor Flags.C
   Flags.Z = (Regs8.A And &HFF) = 0
   Regs8.A = Regs8.A And &HFF
End Sub

Private Sub Add16(ByRef RdestH As Integer, ByRef RdestL As Integer, Reg2H As Integer, Reg2L As Integer)
Dim regDest, RegSrc As Long

   Flags.N = False
   RegSrc = CLng(Reg2H) * 256 + Reg2L
   regDest = CLng(RdestH) * 256 + RdestL
   Flags.H = (((RegSrc And &HFFF) + (regDest And &HFFF)) And &H1000) > 0
   Flags.P = (((RegSrc And &H7FFF) + (regDest And &H7FFF)) And CLng("&H8000")) > 0
   regDest = regDest + RegSrc
   Flags.C = (regDest And &H10000) > 0
   regDest = (regDest And CLng("&HFFFF"))
   RdestH = (regDest And &HFF00) / 256
   RdestL = regDest And &HFF
End Sub

Private Sub Adc16(RdestH As Integer, RdestL As Integer, RsrcH As Integer, RsrcL As Integer)
Dim Dest As Long
Dim Src As Long
Dim result As Long

   Flags.N = False
   Dest = CLng(RdestH) * 256 + RdestL
   If Flags.C Then Dest = Dest + 1
   Src = CLng(RsrcH) * 256 + RsrcL
   Flags.C = ((Dest + Src) And &H10000) > 0
   Flags.H = Flags.C Xor ((((Dest And &HFFF) + (Src And &HFFF)) And &H1000) > 0)
   Flags.P = Flags.C Xor ((((Dest And &H7FFF) + (Src And &H7FFF)) And CLng("&H8000")) > 0)
   result = (Dest + Src) And CLng("&HFFFF")
   result = result And CLng(&HFFFF)
   Flags.Z = result = 0
   Flags.S = (result And CLng("&H8000")) > 0
   
   RdestH = (result And CLng("&HFF00")) / 256
   RdestL = result And &HFF
End Sub

Private Sub Sub8(reg As Integer)
Dim temp As Integer
   temp = Regs8.A - reg
   
   Flags.N = True                              ' an 8 bit subtract
   Flags.S = (temp And &H80) > 0
   Flags.H = ((Regs8.A And &HF) - (reg And &HF)) < 0
   Flags.P = ((Regs8.A And &H80) And Not (temp And &H80))
   Flags.C = (temp And &H100) > 0
   Regs8.A = temp And &HFF
   Flags.Z = (Regs8.A = 0)
End Sub

Private Sub Sbc8(ByVal reg As Integer)
Dim temp As Integer
   
   temp = Regs8.A - reg
   
   Flags.N = True                              ' an 8 bit subtract
   Flags.S = (temp And &H80) > 0
   Flags.H = ((Regs8.A And &HF) - (reg And &HF)) < 0
   Flags.P = ((Regs8.A And &H80) And Not (temp And &H80))
   If Flags.C Then temp = temp - 1             ' carry flag from previous operation
   
   Flags.C = (temp And &H100) > 0
   Regs8.A = temp And &HFF
   
   Flags.Z = (Regs8.A = 0)
   
End Sub

Private Sub Sbc16(RsrcH As Integer, RsrcL As Integer)
Dim Dest As Long
Dim Src As Long
Dim result As Long
   
   Dest = CLng(Regs8.H) * 256 + Regs8.L
   If Flags.C Then Dest = Dest + 1
   Src = CLng(RsrcH) * 256 + RsrcL
   Flags.C = Src > Dest
   Flags.H = Flags.C Xor (((Dest And &HFFF) - (Src And &HFFF)) < 0)
   Flags.P = Flags.C Xor ((((Dest And &H7FFF) - (Src And &H7FFF)) And CLng("&H8000")) > 0)
   result = (Dest - Src) And CLng("&HFFFF")
   Flags.Z = result = 0
   Flags.S = (result And CLng("&H8000")) > 0
   Flags.N = True
   Regs8.H = (result And CLng("&HFF00")) / 256
   Regs8.L = result And &HFF
End Sub

Private Sub Andd(reg As Integer)
   Regs8.A = Regs8.A And reg
   Regs8.A = Regs8.A And &HFF
   Flags.Z = (Regs8.A = 0)
   Flags.H = True
   Flags.P = Parity(Regs8.A)
   Flags.N = False
   Flags.C = False
End Sub

Private Sub Xorr(reg As Integer)
   Regs8.A = Regs8.A Xor reg
   Regs8.A = Regs8.A And &HFF
   Flags.S = (Regs8.A And &H80) > 0
   Flags.Z = Regs8.A = 0
   Flags.H = False
   Flags.P = Parity(Regs8.A)
   Flags.N = False
   Flags.C = False
End Sub

Private Sub Orr(reg As Integer)
   Regs8.A = Regs8.A Or reg
   Regs8.A = Regs8.A And &HFF
   Flags.S = (Regs8.A And &H80) > 0
   Flags.Z = Regs8.A = 0
   Flags.H = False
   Flags.P = Parity(Regs8.A)
   Flags.N = False
   Flags.C = False
End Sub

Private Sub Cp(reg As Integer)
Dim temp As Integer
   temp = Regs8.A - reg
   Flags.N = True
   Flags.S = ((temp And &H80) > 0)
   Flags.H = ((Regs8.A And &HF) - (reg And &HF)) < 0
   Flags.Z = (temp = 0)
   Flags.P = ((Regs8.A And &H80) And Not (temp And &H80))
   Flags.C = (temp And &H100) > 0
End Sub

Private Sub DAA()
Dim LoNibl, HiNibl As Integer
Dim Adder As Integer
   Adder = 0
   LoNibl = Regs8.A And &HF
   HiNibl = (Regs8.A And &HF0) / 16
   If Not Flags.N Then
      If Not Flags.C Then                                               ' N C Hi  H Lo  add
         If (HiNibl < 10) And (Not Flags.H) And (LoNibl < 10) Then      ' 0 0 0-9 0 0-9 00
            Flags.C = False
         ElseIf (HiNibl < 9) And (Not Flags.H) And (LoNibl > 9) Then    ' 0 0 0-8 0 A-F 06
            Adder = &H6
            Flags.C = False
         ElseIf (HiNibl < 10) And Flags.H And (LoNibl < 4) Then         ' 0 0 0-9 1 0-3 06
            Adder = &H6
            Flags.C = False
         ElseIf (HiNibl > 9) And (Not Flags.H) And (LoNibl < 10) Then  ' 0 0 A-F 0 0-9 60
            Adder = &H60
            Flags.C = True
         ElseIf (HiNibl > 8) And (Not Flags.H) And (LoNibl > 9) Then    ' 0 0 9-F 1 A-F 66
            Adder = &H66
            Flags.C = True
         ElseIf (HiNibl > 9) And (Flags.H) And (LoNibl < 4) Then        ' 0 0 A-F 1 0-3 66
            Adder = &H66
            Flags.C = True
         End If
      Else
         If (HiNibl < 3) And (Not Flags.H) And (LoNibl < 10) Then       ' 0 1 0-2 0 0-9 60
            Adder = &H60
            Flags.C = True
         ElseIf (HiNibl < 3) And (Not Flags.H) And (LoNibl > 9) Then    ' 0 1 0-2 0 A-F 66
            Adder = &H66
            Flags.C = True
         ElseIf (HiNibl < 4) And (Flags.H) And (LoNibl < 4) Then        ' 0 1 0-3 1 0-3 66
            Adder = &H66
            Flags.C = True
         End If
      End If
   End If
   If Flags.N Then
      If Not Flags.C Then                                               ' N C Hi  H Lo  sub
         If (HiNibl < 10) And (Not Flags.H) And (LoNibl < 10) Then      ' 1 0 0-9 0 0-9
            Adder = 0
            Flags.C = False
         ElseIf (HiNibl < 9) And (Flags.H) And (LoNibl > 5) Then        ' 1 0 0-8 1 6-f
            Adder = &HFA
            Flags.C = False
         End If
      Else
         If (HiNibl > 6) And (Not Flags.H) And (LoNibl < 10) Then       ' 1 1 7-F 0 0-9
            Adder = &HA0
            Flags.C = True
         ElseIf (HiNibl > 5) And (Flags.H) And (LoNibl > 5) Then        ' 1 1 6-F 1 6-f
            Adder = &H9A
            Flags.C = True
         End If
      End If
   End If
   Regs8.A = (Regs8.A + Adder) And &HFF
   Flags.Z = Regs8.A = 0
   Flags.S = (Regs8.A And &H80)
   Flags.P = Parity(Regs8.A)
End Sub

Private Sub swap(int1 As Integer, int2 As Integer)
Dim temp As Integer
   temp = int1
   int1 = int2
   int2 = temp
End Sub

Public Sub Emulate()
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim Temp1, Temp2 As Long
   
   Halted = False
   OpCode = NextByte
   
   Select Case OpCode
      Case &H0:                                           ' nop
      Case &H1: data1 = NextByte                          ' ld bc,nn
                Data2 = NextByte
                Regs8.C = data1
                Regs8.B = Data2
      Case &H2: Call Write8(Regs8.B, Regs8.C, Regs8.A)    ' ld (bc),a
      Case &H3: Call Incr16(Regs8.B, Regs8.C)             ' inc bc
      Case &H4: Regs8.B = Incr8(Regs8.B)                      ' inc b
      Case &H5: Regs8.B = Decr8(Regs8.B)                      ' dec b
      Case &H6: data1 = NextByte                          ' ld b,n
                Regs8.B = data1
      Case &H7: Call Rlc(8, 0)                            ' rlca
      Case &H8: Temp1 = AltRegs8.A                        ' ex af,af'
                Temp2 = AltRegs8.F
                AltRegs8.A = Regs8.A
                AltRegs8.F = Regs8.F
                Regs8.A = Temp1
                Regs8.F = Temp2
                Flags.S = (Regs8.F And &H80) > 0
                Flags.Z = (Regs8.F And &H40) > 0
                Flags.H = (Regs8.F And &H10) > 0
                Flags.P = (Regs8.F And &H4) > 0
                Flags.N = (Regs8.F And &H2) > 0
                Flags.C = (Regs8.F And &H1) > 0
      Case &H9: Call Add16(Regs8.H, Regs8.L, Regs8.B, Regs8.C) ' add hl,bc
      Case &HA: Regs8.A = Read8(Regs8.B, Regs8.C)         ' ld a,(bc)
      Case &HB: Call Decr16(Regs8.B, Regs8.C)             ' dec bc
      Case &HC: Regs8.C = Incr8(Regs8.C)                      ' inc c
      Case &HD: Regs8.C = Decr8(Regs8.C)                      ' dec c
      Case &HE: data1 = NextByte                          ' ld c,n
                Regs8.C = data1
      Case &HF: Call Rrc(8, 0)                            ' rrca
      Case &H10: data1 = NextByte                         ' djnz
                 If (data1 And &H80) > 0 Then data1 = data1 - 256
                 Regs8.B = Decr8(Regs8.B)
                 If Not Flags.Z Then Regs16.PC = (Regs16.PC + data1) And CLng("&HFFFF")
      Case &H11: data1 = NextByte                         ' ld de,nn
                 Data2 = NextByte
                 Regs8.E = data1
                 Regs8.D = Data2
      Case &H12: Call Write8(Regs8.D, Regs8.E, Regs8.A)   ' ld (de),a
      Case &H13: Call Incr16(Regs8.D, Regs8.E)            ' inc de
      Case &H14: Regs8.D = Incr8(Regs8.D)                     ' inc d
      Case &H15: Regs8.D = Decr8(Regs8.D)                    ' dec d
      Case &H16: data1 = NextByte                         ' ld d,n
                 Regs8.D = data1
      Case &H17: Call Rl(8, 0)
      Case &H18: data1 = NextByte                         ' jr n
                 If (data1 And &H80) > 0 Then data1 = data1 - 256
                 Regs16.PC = (Regs16.PC + data1) And CLng("&HFFFF")
      Case &H19: Call Add16(Regs8.H, Regs8.L, Regs8.D, Regs8.E) ' add hl,de
      Case &H1A: Regs8.A = Read8(Regs8.D, Regs8.E)        ' ld a,(de)
      Case &H1B: Call Decr16(Regs8.D, Regs8.E)            ' dec de
      Case &H1C: Regs8.E = Incr8(Regs8.E)                     ' inc e
      Case &H1D: Regs8.E = Decr8(Regs8.E)                     ' dec e
      Case &H1E: data1 = NextByte                         ' ld e,n
                  Regs8.E = data1
      Case &H1F: Call Rr(8, 0)
      Case &H20: data1 = NextByte                         ' jr nz,n
                 If (data1 And &H80) > 0 Then data1 = data1 - 256
                 If Not Flags.Z Then Regs16.PC = (Regs16.PC + data1) And CLng("&HFFFF")
      Case &H21: data1 = NextByte                         ' ld hl,nn
                  Data2 = NextByte
                  Regs8.L = data1
                  Regs8.H = Data2
      Case &H22: data1 = NextByte                         ' ld (nn),hl
                 Data2 = NextByte
                 Call Write16(Data2, data1, Regs8.H, Regs8.L)
      Case &H23: Call Incr16(Regs8.H, Regs8.L)            ' inc hl
      Case &H24: Regs8.H = Incr8(Regs8.H)                     ' inc h
      Case &H25: Regs8.H = Decr8(Regs8.H)                     ' dec h
      Case &H26: data1 = NextByte                         ' ld h,n
                 Regs8.H = data1
      Case &H27: Call DAA                                 ' daa
      Case &H28: data1 = NextByte                         ' jr z,n
                  If (data1 And &H80) > 0 Then data1 = data1 - 256
                 If Flags.Z Then Regs16.PC = (Regs16.PC + data1) And CLng("&HFFFF")
     Case &H29: Call Add16(Regs8.H, Regs8.L, Regs8.H, Regs8.L) ' add hl,hl
      Case &H2A: data1 = NextByte                         ' ld hl,(nn)
                 Data2 = NextByte
                 Regs8.L = RAM(CLng(Data2) * 256 + data1)
                 Regs8.H = RAM((CLng(Data2) * 256 + data1 + 1) And CLng("&HFFFF"))
      Case &H2B: Call Decr16(Regs8.H, Regs8.L)            ' dec hl
      Case &H2C: Regs8.L = Incr8(Regs8.L)                     ' inc l
      Case &H2D: Regs8.L = Decr8(Regs8.L)                     ' dec l
      Case &H2E: data1 = NextByte                         ' ld l,n
                 Regs8.L = data1
      Case &H2F: Regs8.A = (Not Regs8.A) And &HFF         ' cpl
                 Flags.H = True
                 Flags.N = True
      Case &H30: data1 = NextByte                         ' jr nc,nn
                 If (data1 And &H80) > 0 Then data1 = data1 - 256
                 If Not Flags.C Then Regs16.PC = (Regs16.PC + data1) And CLng("&HFFFF")
      Case &H31: data1 = NextByte                         ' ld sp,nn
                 Data2 = NextByte
                 Regs16.SP = data1 + CLng(Data2) * 256
      Case &H32: data1 = NextByte                         ' ld (nn),a
                 Data2 = NextByte
                 Call Write8(Data2, data1, Regs8.A)
      Case &H33: Regs16.SP = (Regs16.SP + 1) And CLng("&HFFFF")
      Case &H34: Regs8.M = Read8(Regs8.H, Regs8.L)        ' inc (hl)
                 Flags.H = (Regs8.M And &HF) = &HF
                 Flags.P = (Regs8.M And &H7F) = &H7F
                 Flags.N = False
                 Regs8.M = Incr8(Regs8.M)
                 Call Write8(Regs8.H, Regs8.L, Regs8.M)
      Case &H35: Regs8.M = Read8(Regs8.H, Regs8.L)        ' dec (hl)
                 Regs8.M = Decr8(Regs8.M)
                 Call Write8(Regs8.H, Regs8.L, Regs8.M)
      Case &H36: data1 = NextByte                         ' ld (hl),n
                 Call Write8(Regs8.H, Regs8.L, data1)
      Case &H37: Flags.C = True                           ' scf
                 Flags.H = False
                 Flags.N = False
      Case &H38: data1 = NextByte                         ' jr c,nn
                 If (data1 And &H80) > 0 Then data1 = data1 - 256
                 If Flags.C Then Regs16.PC = (Regs16.PC + data1) And CLng("&HFFFF")
      Case &H39: Call Add16(Regs8.H, Regs8.L, Regs8.S, Regs8.P) ' add hl,sp
      Case &H3A: data1 = NextByte                         ' ld a,(nn)
                 Data2 = NextByte
                 Regs8.A = Read8(Data2, data1)
      Case &H3B: Regs16.SP = (Regs16.SP - 1) And CLng("&HFFFF") ' dec sp
      Case &H3C: Regs8.A = Incr8(Regs8.A)                     ' inc a
      Case &H3D: Regs8.A = Decr8(Regs8.A)                     ' dec a
      Case &H3E: data1 = NextByte                         ' ld a,n
                 Regs8.A = data1
      Case &H3F: Flags.H = Flags.C                        ' ccf
                 Flags.C = Not Flags.C
                 Flags.N = False
      Case &H40: Regs8.B = Regs8.B                        ' ld b,b
      Case &H41: Regs8.B = Regs8.C                        ' ld b,c
      Case &H42: Regs8.B = Regs8.D                        ' ld b,d
      Case &H43: Regs8.B = Regs8.E                        ' ld b,e
      Case &H44: Regs8.B = Regs8.H                        ' ld b,h
      Case &H45: Regs8.B = Regs8.L                        ' ld b,l
      Case &H46: Regs8.B = Read8(Regs8.H, Regs8.L)        ' ld b,(hl)
      Case &H47: Regs8.B = Regs8.A                        ' ld b,a
      Case &H48: Regs8.C = Regs8.B                        ' ld c,b
      Case &H49: Regs8.C = Regs8.C                        ' ld c,c
      Case &H4A: Regs8.C = Regs8.D                        ' ld c,d
      Case &H4B: Regs8.C = Regs8.E                        ' ld c,e
      Case &H4C: Regs8.C = Regs8.H                        ' ld c,h
      Case &H4D: Regs8.C = Regs8.L                        ' ld c,l
      Case &H4E: Regs8.C = Read8(Regs8.H, Regs8.L)        ' ld c,(hl)
      Case &H4F: Regs8.C = Regs8.A                        ' ld c,a
      Case &H50: Regs8.D = Regs8.B                        ' ld d,b
      Case &H51: Regs8.D = Regs8.C                        ' ld d,c
      Case &H52: Regs8.D = Regs8.D                        ' ld d,d
      Case &H53: Regs8.D = Regs8.E                        ' ld d,e
      Case &H54: Regs8.D = Regs8.H                        ' ld d,h
      Case &H55: Regs8.D = Regs8.L                        ' ld d,l
      Case &H56: Regs8.D = Read8(Regs8.H, Regs8.L)        ' ld d,(hl)
      Case &H57: Regs8.D = Regs8.A                        ' ld d,a
      Case &H58: Regs8.E = Regs8.B                        ' ld e,b
      Case &H59: Regs8.E = Regs8.C                        ' ld e,c
      Case &H5A: Regs8.E = Regs8.D                        ' ld e,d
      Case &H5B: Regs8.E = Regs8.E                        ' ld e,e
      Case &H5C: Regs8.E = Regs8.H                        ' ld e,h
      Case &H5D: Regs8.E = Regs8.L                        ' ld e,l
      Case &H5E: Regs8.E = Read8(Regs8.H, Regs8.L)        ' ld e,(hl)
      Case &H5F: Regs8.E = Regs8.A                        ' ld e,a
      Case &H60: Regs8.H = Regs8.B                        ' ld h,b
      Case &H61: Regs8.H = Regs8.C                        ' ld h,c
      Case &H62: Regs8.H = Regs8.D                        ' ld h,d
      Case &H63: Regs8.H = Regs8.E                        ' ld h,e
      Case &H64: Regs8.H = Regs8.H                        ' ld h,h
      Case &H65: Regs8.H = Regs8.L                        ' ld h,l
      Case &H66: Regs8.H = Read8(Regs8.H, Regs8.L)        ' ld h,(hl)
      Case &H67: Regs8.H = Regs8.A                        ' ld h,a
      Case &H68: Regs8.L = Regs8.B                        ' ld l,b
      Case &H69: Regs8.L = Regs8.C                        ' ld l,c
      Case &H6A: Regs8.L = Regs8.D                        ' ld l,d
      Case &H6B: Regs8.L = Regs8.E                        ' ld l,e
      Case &H6C: Regs8.L = Regs8.H                        ' ld l,h
      Case &H6D: Regs8.L = Regs8.L                        ' ld l,l
      Case &H6E: Regs8.L = Read8(Regs8.H, Regs8.L)        ' ld l,(hl)
      Case &H6F: Regs8.L = Regs8.A                        ' ld l,a
      Case &H70: Call Write8(Regs8.H, Regs8.L, Regs8.B)   ' ld (hl),b
      Case &H71: Call Write8(Regs8.H, Regs8.L, Regs8.C)   ' ld (hl),c
      Case &H72: Call Write8(Regs8.H, Regs8.L, Regs8.D)   ' ld (hl),d
      Case &H73: Call Write8(Regs8.H, Regs8.L, Regs8.E)   ' ld (hl),e
      Case &H74: Call Write8(Regs8.H, Regs8.L, Regs8.H)   ' ld (hl),h
      Case &H75: Call Write8(Regs8.H, Regs8.L, Regs8.L)   ' ld (hl),l
      Case &H76: Regs16.PC = (Regs16.PC - 1) And &H10FFFF ' halt
           Halted = True
      Case &H77: Call Write8(Regs8.H, Regs8.L, Regs8.A)   ' ld (hl),a
      Case &H78: Regs8.A = Regs8.B                        ' ld a,b
      Case &H79: Regs8.A = Regs8.C                        ' ld a,c
      Case &H7A: Regs8.A = Regs8.D                        ' ld a,d
      Case &H7B: Regs8.A = Regs8.E                        ' ld a,e
      Case &H7C: Regs8.A = Regs8.H                        ' ld a,h
      Case &H7D: Regs8.A = Regs8.L                        ' ld a,l
      Case &H7E: Regs8.A = Read8(Regs8.H, Regs8.L)        ' ld a,(hl)
      Case &H7F: Regs8.A = Regs8.A                        ' ld a,a
      Case &H80: Call Add8(Regs8.B)                       ' add a,b
      Case &H81: Call Add8(Regs8.C)                       ' add a,c
      Case &H82: Call Add8(Regs8.D)                       ' add a,d
      Case &H83: Call Add8(Regs8.E)                       ' add a,e
      Case &H84: Call Add8(Regs8.H)                       ' add a,h
      Case &H85: Call Add8(Regs8.L)                       ' add a,l
      Case &H86: data1 = Read8(Regs8.H, Regs8.L)          ' add a,(hl)
                 Call Add8(data1)
      Case &H87: Call Add8(Regs8.A)                       ' add a,a
      Case &H88: Call Adc8(Regs8.B)                       ' adc a,b
      Case &H89: Call Adc8(Regs8.C)                       ' adc a,c
      Case &H8A: Call Adc8(Regs8.D)                       ' adc a,d
      Case &H8B: Call Adc8(Regs8.E)                       ' adc a,e
      Case &H8C: Call Adc8(Regs8.H)                       ' adc a,h
      Case &H8D: Call Adc8(Regs8.L)                       ' adc a,l
      Case &H8E: data1 = Read8(Regs8.H, Regs8.L)          ' adc a,(hl)
                 Call Adc8(data1)
      Case &H8F: Call Adc8(Regs8.A)                       ' adc a,a
      Case &H90: Call Sub8(Regs8.B)                       ' sub a,b
      Case &H91: Call Sub8(Regs8.C)                       ' sub a,c
      Case &H92: Call Sub8(Regs8.D)                       ' sub a,d
      Case &H93: Call Sub8(Regs8.E)                       ' sub a,e
      Case &H94: Call Sub8(Regs8.H)                       ' sub a,h
      Case &H95: Call Sub8(Regs8.L)                       ' sub a,l
      Case &H96: data1 = Read8(Regs8.H, Regs8.L)          ' sub a,(hl)
                  Call Sub8(data1)
      Case &H97: Call Sub8(Regs8.A)                       ' sub a,a
      Case &H98: Call Sbc8(Regs8.B)                       ' sbc a,b
      Case &H99: Call Sbc8(Regs8.C)                       ' sbc a,c
      Case &H9A: Call Sbc8(Regs8.D)                       ' sbc a,c
      Case &H9B: Call Sbc8(Regs8.E)                       ' sbc a,e
      Case &H9C: Call Sbc8(Regs8.H)                       ' sbc a,h
      Case &H9D: Call Sbc8(Regs8.L)                       ' sbc a,l
      Case &H9E: data1 = Read8(Regs8.H, Regs8.L)          ' sbc a,(hl)
                  Call Sbc8(data1)
      Case &H9F: Call Sbc8(Regs8.A)                       ' sbc a,a
      Case &HA0: Call Andd(Regs8.B)                       ' and b
      Case &HA1: Call Andd(Regs8.C)                       ' and c
      Case &HA2: Call Andd(Regs8.D)                       ' and d
      Case &HA3: Call Andd(Regs8.E)                       ' and e
      Case &HA4: Call Andd(Regs8.H)                       ' and h
      Case &HA5: Call Andd(Regs8.L)                       ' and l
      Case &HA6: data1 = Read8(Regs8.H, Regs8.L)          ' and (hl)
                  Call Andd(data1)
      Case &HA7: Call Andd(Regs8.A)                       ' and a
      Case &HA8: Call Xorr(Regs8.B)                       ' xor b
      Case &HA9: Call Xorr(Regs8.C)                       ' xor c
      Case &HAA: Call Xorr(Regs8.D)                       ' xor d
      Case &HAB: Call Xorr(Regs8.E)                       ' xor e
      Case &HAC: Call Xorr(Regs8.H)                       ' xor h
      Case &HAD: Call Xorr(Regs8.L)                       ' xor l
      Case &HAE:  data1 = Read8(Regs8.H, Regs8.L)         ' xor (hl)
                  Call Xorr(data1)
      Case &HAF: Call Xorr(Regs8.A)                       ' xor a
      Case &HB0: Call Orr(Regs8.B)                        ' or b
      Case &HB1: Call Orr(Regs8.C)                        ' or c
      Case &HB2: Call Orr(Regs8.D)                        ' or d
      Case &HB3: Call Orr(Regs8.E)                        ' or e
      Case &HB4: Call Orr(Regs8.H)                        ' or h
      Case &HB5: Call Orr(Regs8.L)                        ' or l
      Case &HB6: data1 = Read8(Regs8.H, Regs8.L)          ' or (hl)
                  Call Orr(data1)
      Case &HB7: Call Orr(Regs8.A)                        ' or a
      Case &HB8: Call Cp(Regs8.B)                         ' cp b
      Case &HB9: Call Cp(Regs8.C)                         ' cp c
      Case &HBA: Call Cp(Regs8.D)                         ' cp d
      Case &HBB: Call Cp(Regs8.E)                         ' cp e
      Case &HBC: Call Cp(Regs8.H)                         ' cp h
      Case &HBD: Call Cp(Regs8.L)                         ' cp l
      Case &HBE: data1 = Read8(Regs8.H, Regs8.L)          ' cp (hl)
                  Call Cp(data1)
      Case &HBF: Call Cp(Regs8.A)                         ' cp a
      Case &HC0: If Not Flags.Z Then Regs16.PC = pop      ' ret nz
      Case &HC1: Temp1 = pop                              ' pop bc
                 Regs8.B = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.C = Temp1 And &HFF
      Case &HC2: data1 = NextByte                         ' jp nz,nn
                 Data2 = NextByte
                 If Not Flags.Z Then Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
      Case &HC3: data1 = NextByte                         ' jp nn
                 Data2 = NextByte
                 Regs16.PC = (CLng(Data2) * 256 + data1)
      Case &HC4: data1 = NextByte                         ' call nz,nn
                 Data2 = NextByte
                 If Not Flags.Z Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1)
                 End If
      Case &HC5: Call Push(CLng(Regs8.B) * 256 + Regs8.C) ' push bc
      Case &HC6: data1 = NextByte                         ' add a,n
                 Call Add8(data1)
      Case &HC7: Push (Regs16.PC)                         ' rst 0
                 Regs16.PC = &H0
      Case &HC8: If Flags.Z Then Regs16.PC = pop          ' ret z
      Case &HC9: Regs16.PC = pop                          ' ret
      Case &HCA: data1 = NextByte                         ' jp z,nn
                 Data2 = NextByte
                 If Flags.Z Then Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
      Case &HCB: Call ExeCB                               ' cb extended operations
      
      Case &HCC: data1 = NextByte                         ' call z,nn
                 Data2 = NextByte
                 If Flags.Z Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1)
                 End If
      Case &HCD: data1 = NextByte                         ' call nn
                 Data2 = NextByte
                 Call Push(Regs16.PC)
                 Regs16.PC = (CLng(Data2) * 256 + data1)
      Case &HCE: data1 = NextByte                         ' adc a,n
                  Call Adc8(data1)
      Case &HCF: Call Push(Regs16.PC)                     ' rst 8
                 Regs16.PC = &H8
      Case &HD0: If Not Flags.C Then Regs16.PC = pop      ' ret nc
      Case &HD1: Temp1 = pop                              ' pop de
                 Regs8.D = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.E = Temp1 And &HFF
      Case &HD2: data1 = NextByte                         ' jp nc,nn
                 Data2 = NextByte
                 If Not Flags.C Then Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
      Case &HD3: data1 = NextByte                         ' out (n),a
      Case &HD4: data1 = NextByte                         ' call nc,nn
                 Data2 = NextByte
                 If Not Flags.C Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1)
                 End If
      Case &HD5: Call Push(CLng(Regs8.D) * 256 + Regs8.E) ' push de
      Case &HD6: data1 = NextByte                         ' sub a,n
                  Call Sub8(data1)
      Case &HD7: Call Push(Regs16.PC)                     ' rst 16
                 Regs16.PC = &H16
      Case &HD8: If Flags.C Then Regs16.PC = pop          ' ret c
      Case &HD9: Call swap(Regs8.B, AltRegs8.B)            ' exx
                 Call swap(Regs8.C, AltRegs8.C)
                 Call swap(Regs8.D, AltRegs8.D)
                 Call swap(Regs8.E, AltRegs8.E)
                 Call swap(Regs8.H, AltRegs8.H)
                 Call swap(Regs8.L, AltRegs8.L)
      Case &HDA: data1 = NextByte                         ' jp c,nn
                 Data2 = NextByte
                 If Flags.C Then Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
      Case &HDB: data1 = NextByte                         ' in a,(n)
      Case &HDC: data1 = NextByte                         ' call c,nn
                 Data2 = NextByte
                 If Flags.C Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
                 End If
      Case &HDD: Call ExeDDFD(&H100)                      ' dd extended opcodes
      Case &HDE: data1 = NextByte                         ' sbc a,n
                  Call Sbc8(data1)
      Case &HDF: Call Push(Regs16.PC)                     ' rst 24
                 Regs16.PC = &H24
      Case &HE0: If Not Flags.P Then Regs16.PC = pop      ' ret po
      Case &HE1: Temp1 = pop                              ' pop hl
                 Regs8.H = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.L = Temp1 And &HFF
      Case &HE2: data1 = NextByte                         ' jp po,nn
                 Data2 = NextByte
                 If Not Flags.P Then Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
      Case &HE3: Temp1 = Regs8.H                          ' ex (sp),hl
                 Temp2 = Regs8.L
                 Regs8.L = RAM(Regs16.SP)
                 Regs8.H = RAM((Regs16.SP + 1)) And CLng("&HFFFF")
                 RAM(Regs16.SP) = Temp2
                 RAM((Regs16.SP + 1) And CLng("&HFFFF")) = Temp1
      Case &HE4: data1 = NextByte                         ' call po,nn
                 Data2 = NextByte
                 If Not Flags.P Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
                 End If
      Case &HE5: Call Push(CLng(Regs8.H) * 256 + Regs8.L) ' push hl
      Case &HE6: data1 = NextByte                         ' and n
                  Call Andd(data1)
      Case &HE7: Call Push(Regs16.PC)                     ' rst 32
                 Regs16.PC = &H32
      Case &HE8: If Flags.P Then Regs16.PC = pop          ' ret pe
      Case &HE9: Regs16.PC = Regs8.H * 265 + Regs8.L     ' jp (hl)
      Case &HEA: data1 = NextByte                        ' jp pe,nn
                 Data2 = NextByte
                 If Flags.P Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
                 End If
      Case &HEB: data1 = Regs8.D                          ' ex de,hl
                 Data2 = Regs8.E
                 Regs8.D = Regs8.H
                 Regs8.E = Regs8.L
                 Regs8.H = data1
                 Regs8.L = Data2
      Case &HEC: data1 = NextByte                         ' call pe,nn
                 Data2 = NextByte
                 If Not Flags.P Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
                 End If
      Case &HED: Call ExeED                               ' ed entended opcodes
      
      Case &HEE: data1 = NextByte                         ' xor n
                  Call Xorr(data1)
      Case &HEF: Call Push(Regs16.PC)                     ' rst 40
                 Regs16.PC = &H40
      Case &HF0: If Not Flags.S Then Regs16.PC = pop      ' ret p
      Case &HF1:
                 Temp1 = pop                              ' pop af
                 Regs8.A = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.F = Temp1 And &HFF
      Case &HF2: data1 = NextByte                         ' jp p,nn
                 Data2 = NextByte
                 If Not Flags.S Then Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
      Case &HF3:                                          ' di
      Case &HF4: data1 = NextByte                         ' call p,nn
                 Data2 = NextByte
                 If Not Flags.S Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
                 End If
      Case &HF5:
                 Call Push(CLng(Regs8.A) * 256 + Regs8.F) ' push af
      Case &HF6: data1 = NextByte                         ' or n
                  Call Orr(data1)
      Case &HF7: Call Push(Regs16.PC)                     ' rst 48
                 Regs16.PC = &H48
      Case &HF8: If Flags.S Then Regs16.PC = pop          ' ret m
      Case &HF9: Regs16.SP = Regs8.H * CLng(256) + Regs8.L      ' ld sp,hl
      Case &HFA: data1 = NextByte                         ' jp m,nn
                 Data2 = NextByte
                 If Flags.S Then Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
      Case &HFB:                                          ' ei
      Case &HFC: data1 = NextByte                         ' call m,nn
                 Data2 = NextByte
                 If Flags.S Then
                    Call Push(Regs16.PC)
                    Regs16.PC = (CLng(Data2) * 256 + data1) And CLng("&HFFFF")
                 End If
      Case &HFD: Call ExeDDFD(&H200)                      ' fd extended opcodes
      
      Case &HFE: data1 = NextByte                         ' cp n
                  Call Cp(data1)
      Case &HFF: Call Push(Regs16.PC)                     ' rst 56
                 Regs16.PC = &H56
   End Select
   
   Regs8.R = (Regs8.R + 1) And &H7F

End Sub

Private Function Parity(RegVal As Integer) As Boolean
   Parity = (((RegVal And &H80) \ &H80) + ((RegVal And &H40) \ &H40) + ((RegVal And &H20) \ &H20) + ((RegVal And &H10) \ &H10) + _
              ((RegVal And 8) \ 8) + ((RegVal And 4) \ 4) + ((RegVal And 2) \ 2) + ((RegVal And &H1))) Mod 2 = 0
End Function

Private Sub Rlc(reg As Integer, Offset As Integer)
Dim RegVal As Integer
   
   Select Case reg
      Case 0: RegVal = Regs8.B   ' rlc b
      Case 1: RegVal = Regs8.C   ' rlc c
      Case 2: RegVal = Regs8.D   ' rlc d
      Case 3: RegVal = Regs8.E   ' rlc e
      Case 4: RegVal = Regs8.H   ' rlc h
      Case 5: RegVal = Regs8.L   ' rlc l
      Case 6: RegVal = RAM(CLng(Regs8.H) * 256 + Regs8.L)  ' rlc (hl)
      Case 7: RegVal = Regs8.A   ' rlc a
      Case 8: RegVal = Regs8.A   ' rlca
      Case &H100: RegVal = RAM(Regs16.IX + Offset)  ' rlc (ix+d)
      Case &H200: RegVal = RAM(Regs16.IY + Offset)  ' RLC (IY+D)
   End Select
   RegVal = RegVal * 2
   If (RegVal And &H100) > 0 Then RegVal = RegVal + 1
   Select Case reg
      Case 0: Regs8.B = RegVal
      Case 1: Regs8.C = RegVal
      Case 2: Regs8.D = RegVal
      Case 3: Regs8.E = RegVal
      Case 4: Regs8.H = RegVal
      Case 5: Regs8.L = RegVal
      Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RegVal
      Case 7: Regs8.A = RegVal
      Case 8: Regs8.A = RegVal
      Case &H100: RAM(CLng(Regs16.IX + Offset)) = RegVal
      Case &H200: RAM(Regs16.IY + Offset) = RegVal
   End Select
    
   Flags.H = False
   Flags.N = False
   If reg <> 8 Then  ' rlca affects the flags differetly than the other rlc r instructions
      Flags.P = Parity(RegVal)
      Flags.S = (RegVal And &H80) > 0
      Flags.Z = RegVal = 0
   End If
   Flags.C = (RegVal And &H100) > 0
End Sub

Private Sub Rrc(reg As Integer, Offset As Integer)
Dim RegVal As Integer
Dim remainder As Integer

   Select Case reg
      Case 0: RegVal = Regs8.B   ' rrc b
      Case 1: RegVal = Regs8.C   ' rrc c
      Case 2: RegVal = Regs8.D   ' rrc d
      Case 3: RegVal = Regs8.E   ' rrc e
      Case 4: RegVal = Regs8.H   ' rrc h
      Case 5: RegVal = Regs8.L   ' rrc l
      Case 6: RegVal = RAM(CLng(Regs8.H) * 256 + Regs8.L)   ' rrc (hl)
      Case 7: RegVal = Regs8.A   ' rrc a
      Case 8: RegVal = Regs8.A   ' rrca
      Case &H100: RegVal = RAM(Regs16.IX + Offset)
      Case &H200: RegVal = RAM(Regs16.IY + Offset)
   End Select
   remainder = RegVal Mod 2
   RegVal = (RegVal \ 2) + (remainder * 128)
   Select Case reg
      Case 0: Regs8.B = RegVal
      Case 1: Regs8.C = RegVal
      Case 2: Regs8.D = RegVal
      Case 3: Regs8.E = RegVal
      Case 4: Regs8.H = RegVal
      Case 5: Regs8.L = RegVal
      Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RegVal
      Case 7: Regs8.A = RegVal
      Case 8: Regs8.A = RegVal
      Case &H100: RAM(Regs16.IX + Offset) = RegVal
      Case &H200: RAM(Regs16.IY + Offset) = RegVal
   End Select
    
       
   Flags.H = False
   Flags.N = False
   If reg <> 8 Then  ' rrca affects the flags differetly than the other rrc r instructions
      Flags.S = (RegVal And &H80) > 0
      Flags.P = Parity(RegVal)
      Flags.N = False
      Flags.Z = False
   End If
   Flags.C = remainder > 0
End Sub

Private Sub Rl(reg As Integer, Offset As Integer)
Dim RegVal As Integer

   Select Case reg
      Case 0: RegVal = Regs8.B
      Case 1: RegVal = Regs8.C
      Case 2: RegVal = Regs8.D
      Case 3: RegVal = Regs8.E
      Case 4: RegVal = Regs8.H
      Case 5: RegVal = Regs8.L
      Case 6: RegVal = RAM(CLng(Regs8.H) * 256 + Regs8.L)
      Case 7: RegVal = Regs8.A
      Case 8: RegVal = Regs8.A
      Case &H100: RegVal = RAM(Regs16.IX + Offset)
      Case &H200: RegVal = RAM(Regs16.IY + Offset)
   End Select
   RegVal = RegVal * 2
   If Flags.C Then RegVal = RegVal + 1
   Select Case reg
      Case 0: Regs8.B = RegVal
      Case 1: Regs8.C = RegVal
      Case 2: Regs8.D = RegVal
      Case 3: Regs8.E = RegVal
      Case 4: Regs8.H = RegVal
      Case 5: Regs8.L = RegVal
      Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RegVal
      Case 7: Regs8.A = RegVal
      Case 8: Regs8.A = RegVal
      Case &H100: RAM(Regs16.IX + Offset) = RegVal
      Case &H200: RAM(Regs16.IY + Offset) = RegVal
   End Select
   
   If reg <> 8 Then  ' rla affects the flags differetly than the other rl r instructions
      Flags.S = (RegVal And &H80) > 0
      Flags.Z = (RegVal And &HFF) = 0
      Flags.P = Parity(RegVal)
   End If
   Flags.H = False
   Flags.N = False
   Flags.C = (RegVal And &H100) > 0
End Sub

Private Sub Rr(reg As Integer, Offset As Integer)
Dim RegVal As Integer
Dim remainder As Integer
   Select Case reg
      Case 0: RegVal = Regs8.B
      Case 1: RegVal = Regs8.C
      Case 2: RegVal = Regs8.D
      Case 3: RegVal = Regs8.E
      Case 4: RegVal = Regs8.H
      Case 5: RegVal = Regs8.L
      Case 6: RegVal = RAM(CLng(Regs8.H) * 256 + Regs8.L)
      Case 7: RegVal = Regs8.A
      Case 8: RegVal = Regs8.A
      Case &H100: RegVal = RAM(Regs16.IX + Offset)
      Case &H200: RegVal = RAM(Regs16.IY + Offset)
   End Select
   remainder = RegVal Mod 2
   RegVal = RegVal \ 2
   If Flags.C Then RegVal = RegVal + 128
   Select Case reg
      Case 0: Regs8.B = RegVal
      Case 1: Regs8.C = RegVal
      Case 2: Regs8.D = RegVal
      Case 3: Regs8.E = RegVal
      Case 4: Regs8.H = RegVal
      Case 5: Regs8.L = RegVal
      Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RegVal
      Case 7: Regs8.A = RegVal
      Case 8: Regs8.A = RegVal
      Case &H100: RAM(Regs16.IX + Offset) = RegVal
      Case &H200: RAM(Regs16.IY + Offset) = RegVal
   End Select
   
   If reg <> 8 Then  ' rra affects the flags differetly than the other rr r instructions
      Flags.S = (RegVal And &H80) > 0
      Flags.Z = (RegVal And &HFF) = 0
      Flags.P = Parity(RegVal)
   End If
   Flags.H = False
   Flags.N = False
   Flags.C = remainder > 0
End Sub

Private Sub Sla(reg As Integer, Offset As Integer)
Dim RegVal As Integer

   Select Case reg
      Case 0: RegVal = Regs8.B
      Case 1: RegVal = Regs8.C
      Case 2: RegVal = Regs8.D
      Case 3: RegVal = Regs8.E
      Case 4: RegVal = Regs8.H
      Case 5: RegVal = Regs8.L
      Case 6: RegVal = RAM(CLng(Regs8.H) * 256 + Regs8.L)
      Case 7: RegVal = Regs8.A
      Case 8: RegVal = Regs8.A
      Case &H100: RegVal = RAM(Regs16.IX + Offset)
      Case &H200: RegVal = RAM(Regs16.IY + Offset)
   End Select
   RegVal = RegVal * 2
   Flags.C = (RegVal And &H100) > 0
   RegVal = RegVal And &HFF
   
   Select Case reg
      Case 0: Regs8.B = RegVal
      Case 1: Regs8.C = RegVal
      Case 2: Regs8.D = RegVal
      Case 3: Regs8.E = RegVal
      Case 4: Regs8.H = RegVal
      Case 5: Regs8.L = RegVal
      Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RegVal
      Case 7: Regs8.A = RegVal
      Case &H100: RAM(Regs16.IX + Offset) = RegVal
      Case &H200: RAM(Regs16.IY + Offset) = RegVal
   End Select
   
   Flags.S = (RegVal And &H80) > 0
   Flags.Z = (RegVal = 0)
   Flags.H = False
   Flags.N = False
   Flags.P = Parity(RegVal)
   
End Sub

Private Sub Sra(reg As Integer, Offset As Integer)
Dim RegVal As Integer
Dim remainder As Integer

Select Case reg
      Case 0: RegVal = Regs8.B
      Case 1: RegVal = Regs8.C
      Case 2: RegVal = Regs8.D
      Case 3: RegVal = Regs8.E
      Case 4: RegVal = Regs8.H
      Case 5: RegVal = Regs8.L
      Case 6: RegVal = RAM(CLng(Regs8.H) * 256 + Regs8.L)
      Case 7: RegVal = Regs8.A
      Case &H100: RegVal = RAM(Regs16.IX + Offset)
      Case &H200: RegVal = RAM(Regs16.IY + Offset)
   End Select
   remainder = RegVal Mod 2
   RegVal = RegVal \ 2
   If (RegVal And &H40) Then RegVal = RegVal + 128
   Select Case reg
      Case 0: Regs8.B = RegVal
      Case 1: Regs8.C = RegVal
      Case 2: Regs8.D = RegVal
      Case 3: Regs8.E = RegVal
      Case 4: Regs8.H = RegVal
      Case 5: Regs8.L = RegVal
      Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RegVal
      Case 7: Regs8.A = RegVal
      Case &H100: RAM(Regs16.IX + Offset) = RegVal
      Case &H200: RAM(Regs16.IY + Offset) = RegVal
   End Select
    
    Flags.S = (RegVal And &H80) > 0
    Flags.Z = (RegVal And &HFF) = 0
    Flags.H = False
    Flags.N = False
    Flags.P = Parity(RegVal)
    Flags.C = remainder > 0


End Sub

Private Sub Sll(reg As Integer)

End Sub

Private Sub Srl(reg As Integer)
Dim RegVal As Integer
Dim remainder As Integer

   Select Case reg
      Case 0: RegVal = Regs8.B
      Case 1: RegVal = Regs8.C
      Case 2: RegVal = Regs8.D
      Case 3: RegVal = Regs8.E
      Case 4: RegVal = Regs8.H
      Case 5: RegVal = Regs8.L
      Case 6: RegVal = RAM(CLng(Regs8.H) * 256 + Regs8.L)
      Case 7: RegVal = Regs8.A
   End Select
   remainder = RegVal Mod 2
   RegVal = RegVal \ 2
   Select Case reg
      Case 0: Regs8.B = RegVal
      Case 1: Regs8.C = RegVal
      Case 2: Regs8.D = RegVal
      Case 3: Regs8.E = RegVal
      Case 4: Regs8.H = RegVal
      Case 5: Regs8.L = RegVal
      Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RegVal
      Case 7: Regs8.A = RegVal
      Case &H100: RAM(Regs16.IX) = RegVal
      Case &H200: RAM(Regs16.IY) = RegVal
   End Select
    
    Flags.S = (RegVal And &H80) > 0
    Flags.Z = (RegVal And &HFF) = 0
    Flags.H = False
    Flags.N = False
    Flags.P = Parity(RegVal)
    Flags.C = remainder > 0

End Sub

Private Sub ExeCB()
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim dat2 As Integer
   
   OpCode = NextByte
   data1 = (OpCode And &HC0) / 64
   Select Case data1
      Case 0: Data2 = ((OpCode And &H38) / 8)
              Data3 = OpCode And &H7
              Select Case Data2
                 Case 0: Call Rlc(Data3, 0)                         ' rlc
                 Case 1: Call Rrc(Data3, 0)                         ' rrc
                 Case 2: Call Rl(Data3, 0)                          ' rl
                 Case 3: Call Rr(Data3, 0)                          ' rr
                 Case 4: Call Sla(Data3, 0)                         ' sla
                 Case 5: Call Sra(Data3, 0)                         ' sra
                 Case 6: Call Sll(Data3)                            ' sll
                 Case 7: Call Srl(Data3)                            ' srl
              End Select
      Case 1: Data2 = (OpCode And &H38) / 8
              Data3 = OpCode And &H7
              Select Case Data3                                     ' bit x,n
                 Case 0: Flags.Z = (Regs8.B And (2 ^ Data2)) = 0
                 Case 1: Flags.Z = (Regs8.C And (2 ^ Data2)) = 0
                 Case 2: Flags.Z = (Regs8.D And (2 ^ Data2)) = 0
                 Case 3: Flags.Z = (Regs8.E And (2 ^ Data2)) = 0
                 Case 4: Flags.Z = (Regs8.H And (2 ^ Data2)) = 0
                 Case 5: Flags.Z = (Regs8.L And (2 ^ Data2)) = 0
                 Case 6: Flags.Z = (RAM(CLng(Regs8.H) * 256 + Regs8.L) = RAM(CLng(Regs8.H) * 256 + Regs8.L) And (2 ^ Data2)) = 0
                 Case 7: Flags.Z = (Regs8.A And (2 ^ Data2)) = 0
            End Select
            Flags.H = True
            Flags.N = False
      Case 2: Data2 = (OpCode And &H38) / 8
              Data3 = OpCode And &H7
              Select Case Data3                                     ' RES x,n
                 Case 0: Regs8.B = Regs8.B And (&HFF - (2 ^ Data2))
                 Case 1: Regs8.C = Regs8.C And (&HFF - (2 ^ Data2))
                 Case 2: Regs8.D = Regs8.D And (&HFF - (2 ^ Data2))
                 Case 3: Regs8.E = Regs8.E And (&HFF - (2 ^ Data2))
                 Case 4: Regs8.H = Regs8.H And (&HFF - (2 ^ Data2))
                 Case 5: Regs8.L = Regs8.L And (&HFF - (2 ^ Data2))
                 Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RAM(CLng(Regs8.H) * 256 + Regs8.L) And (&HFF - (2 ^ Data2))
                 Case 7: Regs8.A = Regs8.A And (&HFF - (2 ^ Data2))
              End Select
      Case 3: Data2 = (OpCode And &H38) / 8
              Data3 = OpCode And &H7
              Select Case Data3                                     ' SET x,n
                 Case 0: Regs8.B = Regs8.B Or (2 ^ Data2)
                 Case 1: Regs8.C = Regs8.C Or (2 ^ Data2)
                 Case 2: Regs8.D = Regs8.D Or (2 ^ Data2)
                 Case 3: Regs8.E = Regs8.E Or (2 ^ Data2)
                 Case 4: Regs8.H = Regs8.H Or (2 ^ Data2)
                 Case 5: Regs8.L = Regs8.L Or (2 ^ Data2)
                 Case 6: RAM(CLng(Regs8.H) * 256 + Regs8.L) = RAM(CLng(Regs8.H) * 256 + Regs8.L) Or (2 ^ Data2)
                 Case 7: Regs8.A = Regs8.A Or (2 ^ Data2)
              End Select

   End Select
End Sub

Private Sub ExeDDFDCB(RegSel As Integer)
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim Temp1 As Long
Dim Temp2 As Long
Dim Temp3 As Long
Dim workReg As Integer
    
   data1 = NextByte
   OpCode = NextByte
   If RegSel = 0 Then
      workReg = RAM(Regs16.IX + data1)
   Else
      workReg = RAM(Regs16.IY + data1)
   End If
   Select Case OpCode
      Case &H6: Call Rlc(RegSel, data1)                                   ' rlc (IXIY+d)
      Case &HE: Call Rrc(RegSel, data1)                                   ' rrc (IXIY+d)
      Case &H16: Call Rl(RegSel, data1)                                   ' rl (IXIY+d)
      Case &H1E: Call Rr(RegSel, data1)                                   ' rr (IXIY+d)
      Case &H26: Call Sla(RegSel, data1)                                  ' sla (IXIY+d)
      Case &H2E: Call Sra(RegSel, data1)                                  ' sra (IXIY+d)
      Case &H46, &H4E, &H56, &H5E, &H66, &H6E, &H76, &H7E:                ' bit n,(IXIY+d)
                 Flags.Z = (RAM(workReg) = RAM(workReg) And (2 ^ (OpCode And &H7))) = 0
            Flags.H = True
            Flags.N = False
      Case &H86, &H8E, &H96, &H9E, &HA6, &HAE, &HB6, &HBE:                ' res n,(IXIY+d)
                 RAM(workReg) = RAM(workReg) And (&HFF - (2 ^ (OpCode And &H7)))
      Case &HC6, &HCE, &HD6, &HDE, &HE6, &HEE, &HF6, &HFE:                ' set n,(IXIY+d)
                 RAM(workReg) = RAM(workReg) Or (2 ^ (OpCode And &H7))
   End Select
   If RegSel = 0 Then
      RAM(Regs16.IX + data1) = workReg
   Else
      RAM(Regs16.IY + data1) = workReg
   End If

End Sub

Private Sub ExeDDFD(RegSel As Integer)
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim Temp1 As Long
Dim Temp2 As Long
Dim Temp3 As Long
Dim workReg As Long
Dim workHi As Integer
Dim workLo As Integer

   If RegSel = &H100 Then
      workReg = Regs16.IX
   Else
      workReg = Regs16.IY
   End If
   workHi = workReg \ 256
   workLo = workReg Mod 256
   OpCode = NextByte
   Select Case OpCode
      Case &H9: Call Add16(workHi, workLo, Regs8.B, Regs8.C)      ' add IXIY,bc
                workReg = workHi * CLng(256) + workLo
      Case &H19: Call Add16(workHi, workLo, Regs8.D, Regs8.E)     ' add IXIY,de
                 workReg = workHi * CLng(256) + workLo
      Case &H21: data1 = NextByte                                 ' ld IXIY,nn
                 Data2 = NextByte
                 workReg = (Data2 * CLng(256) + data1)
      Case &H22: data1 = NextByte                                 ' ld (nn),IXIY
                 Data2 = NextByte
                 Call Write16(Data2, data1, workHi, workLo)
      Case &H23: Call Incr16(workHi, workLo)                      ' inc IXIY
                 workReg = workHi * CLng(256) + workLo
      Case &H29: Call Add16(workHi, workLo, workHi, workLo)       ' add IXIY,IXIY
                 workReg = workHi * CLng(256) + workLo
      Case &H2A: data1 = NextByte                                 ' ld hl,(nn)
                 Data2 = NextByte
                 workLo = RAM(CLng(Data2) * 256 + data1)
                 workHi = RAM((CLng(Data2) * 256 + data1 + 1) And CLng("&HFFFF"))
                 workReg = workHi * CLng(256) + workLo
      Case &H2B: Call Decr16(workHi, workLo)                    ' dec hl
                 workReg = workHi * CLng(256) + workLo
      Case &H34: data1 = NextByte                                ' inc (IXIY+d)
                 Data2 = RAM((workReg + data1) And CLng("&Hffff"))
                 Flags.H = (Data2 And &HF) = &HF
                 Flags.P = (Data2 And &H7F) = &H7F
                 Flags.N = False
                 data1 = (data1 + 1) And &HFF
                 RAM((workReg + data1) And CLng("&Hffff")) = data1
      Case &H35: data1 = NextByte                                 ' dec (IXIY+d)
                 RAM((workReg + data1) And 65535) = (RAM((workReg + data1) And 65535) - 1) And &HFF
      Case &H36: data1 = NextByte
                 Data2 = NextByte
                 RAM((workReg + data1) And 65535) = Data2
      Case &H39: Call Add16(workHi, workLo, Regs16.SP \ 256, Regs16.SP Mod 256) ' add IXIY,sp
                 workReg = workHi * CLng(256) + workLo
      Case &H46: data1 = NextByte                                 ' ld b,(IXIY+d)
                 Regs8.B = RAM((workReg + data1) And 65535)
      Case &H4E: data1 = NextByte                                 ' ld c,(IXIY+d)
                 Regs8.C = RAM((workReg + data1) And 65535)
      Case &H56: data1 = NextByte                                 ' ld d,(IXIY+d)
                 Regs8.D = RAM((workReg + data1) And 65535)
      Case &H5E: data1 = NextByte                                 ' ld e,(IXIY+d)
                 Regs8.E = RAM((workReg + data1) And 65535)
      Case &H66: data1 = NextByte                                 ' ld h,(IXIY+d)
                 Regs8.H = RAM((workReg + data1) And 65535)
      Case &H6E: data1 = NextByte                                 ' ld l,(IXIY+d)
                 Regs8.L = RAM((workReg + data1) And 65535)
      Case &H70: data1 = NextByte                                 ' ld (IXIY+d),b
                 RAM((workReg + data1) And 65535) = Regs8.B
      Case &H71: data1 = NextByte                                 ' ld (IXIY+d),c
                 RAM((workReg + data1) And 65535) = Regs8.C
      Case &H72: data1 = NextByte                                 ' ld (IXIY+d),d
                 RAM((workReg + data1) And 65535) = Regs8.D
      Case &H73: data1 = NextByte                                 ' ld (IXIY+d),e
                 RAM((workReg + data1) And 65535) = Regs8.E
      Case &H74: data1 = NextByte                                 ' ld (IXIY+d),h
                 RAM((workReg + data1) And 65535) = Regs8.H
      Case &H75: data1 = NextByte                                 ' ld (IXIY+d),l
                 RAM((workReg + data1) And 65535) = Regs8.L
      Case &H77: data1 = NextByte                                 ' ld (IXIY+d),a
                 RAM((workReg + data1) And 65535) = Regs8.A
      Case &H7E: data1 = NextByte                                 ' ld a,(IXIY+d)
                 Regs8.A = RAM((workReg + data1) And 65535)
      Case &H86: data1 = NextByte                                 ' add a,(IXIY+d)
                 Data2 = RAM((workReg + data1) And 65535)
                 Call Add8(Data2)
      Case &H8E: data1 = NextByte                                 ' adc a,(IXIY+d)
                 Data2 = RAM((workReg + data1) And 65535)
                 Call Add8(Data2)
      Case &H96: data1 = NextByte                                 ' sub (IXIY+d)
                 Data2 = RAM((workReg + data1) And 65535)
                 Call Sub8(Data2)
      Case &H9E: data1 = NextByte                                 ' sbc (IXIY+d)
                 Data2 = RAM((workReg + data1) And 65535)
                 Call Sbc8(Data2)
      Case &HA6: data1 = NextByte                                 ' and (IXIY+d)
                 Data2 = RAM((workReg + data1) And 65535)
                 Call Andd(Data2)
      
      Case &HAE: data1 = NextByte                                 ' xor a,(IXIY+d)
                 Data2 = RAM((workReg + data1) And 65535)
                 Call Xorr(Data2)
      Case &HB6: data1 = NextByte                                 ' or (IXIY+d)
                 Data2 = RAM((workReg + data1) And 65535)
                 Call Orr(Data2)
      Case &HBE: data1 = NextByte                                 ' cp (IXIY+d)
                 Data2 = RAM((workReg + data1) And 65535)
                 Call Cp(Data2)
      Case &HCB: Call ExeDDFDCB(RegSel)                           ' ED CB Extended opcodes
      Case &HE1: workReg = pop                                    ' pop IXIY
      Case &HE3: data1 = RAM(Regs16.SP)                           ' ex(sp),IXIY
                 Data2 = RAM((Regs16.SP + 1) And 65535)
                 RAM(Regs16.SP) = workLo
                 RAM((Regs16.SP + 1) And 65535) = workHi
                 workLo = data1
                 workHi = Data2
      Case &HE5: Call Push(workReg)                                ' push IXIY
      Case &HE9: Regs16.PC = RAM(workReg) + RAM((workReg + 1) And 65536) * 256  ' jp IXIY
      Case &HF9:
                  Regs16.SP = workReg                               ' ld sp,IXIY
      Case Else 'bad opcode
   End Select
   If RegSel = &H100 Then
      Regs16.IX = workReg
   Else
      Regs16.IY = workReg
   End If
   
End Sub

Private Sub ExeED()
Dim OpCode As Integer
Dim data1 As Integer
Dim Data2 As Integer
Dim Data3 As Integer
Dim Temp1 As Long
Dim Temp2 As Long
Dim Temp3 As Long

   OpCode = NextByte
   Select Case OpCode
      Case &H40:                                                 ' in b,(c)
      Case &H41:                                                 ' out(c),b
      Case &H42: Call Sbc16(Regs8.B, Regs8.C)                    ' sbc hl,bc
      Case &H43: data1 = NextByte                                ' ld (nn),bc
                 Data2 = NextByte
                 Call Write16(data1, Data2, Regs8.B, Regs8.C)
      Case &H44: If Regs8.A = &H80 Then Flags.P = True           ' neg
                 If Regs8.A = 0 Then Flags.C = True
                 Flags.C = (Regs8.A > 0)
                 Regs8.A = (0 - Regs8.A) And &HFF
                 Flags.S = (Regs8.A And &H80) > 0
                 Flags.Z = (Regs8.A = 0)
                 Flags.H = (Regs8.A And &H7F) <> 0
                 Flags.N = True
                 
      Case &H45: Regs16.PC = pop                                 ' retn
      Case &H46:                                                 ' im 0
      Case &H47:
                 Regs8.I = Regs8.A                               ' ld i,a
      Case &H48:                                                 ' in c,(c)
      Case &H49:                                                 ' out (c),c
      Case &H4A: Call Adc16(Regs8.H, Regs8.L, Regs8.B, Regs8.C)  ' adc hl,bc
      Case &H4B: data1 = NextByte                                ' ld bc,(nn)
                 Data2 = NextByte
                 Temp1 = Read16(data1, Data2)
                 Regs8.B = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.C = (Temp1 And &HFF)
      Case &H4D:                                                 ' reti
      Case &H4F: Regs8.R = Regs8.A                               ' ld r,a
      Case &H50:                                                 ' in d,(c)
      Case &H51:                                                 ' out (c),d
      Case &H52: Call Sbc16(Regs8.D, Regs8.E)                    ' sbc hl,de
      Case &H53: data1 = NextByte                                ' ld (nn),de
                 Data2 = NextByte
                 Call Write16(data1, Data2, Regs8.D, Regs8.E)
      Case &H56:                                                 ' im 1
      Case &H57:
                 Regs8.A = Regs8.I                               ' ld a,i
      Case &H58:                                                 ' in e,(c)
      Case &H59:                                                 ' out (c),e
      Case &H5A: Call Adc16(Regs8.H, Regs8.L, Regs8.D, Regs8.E)  ' adc hl,de
      Case &H5B: data1 = NextByte                                ' ld de,(nn)
                 Data2 = NextByte
                 Temp1 = Read16(data1, Data2)
                 Regs8.D = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.E = (Temp1 And &HFF)
      Case &H5E:                                                 ' im 2
      Case &H5F: Regs8.A = Regs8.R                               ' ld a,r
      Case &H60:                                                 ' in h,(c)
      Case &H61:                                                 ' out (c),h
      Case &H62: Call Sbc16(Regs8.H, Regs8.L)                    ' sbc hl,hl
      Case &H67: data1 = Read8(Regs8.H, Regs8.L)                 ' rrd
                 Data2 = data1 Mod 16  ' data2 holds low nibble of (HL)
                 data1 = data1 \ 16    ' data1 holds high nibble of (HL)
                 Data3 = Regs8.A And &HF ' data3 hold low nibble of A
                 Regs8.A = (Regs8.A And &HF0) + Data2 ' low nibble of (HL) in low nibble A
                 data1 = data1 + (Data3 * 16) ' data1 now has high of (hl) in low and low of A in high
                 Call Write8(Regs8.H, Regs8.L, data1)
                 Flags.S = (Regs8.A And &H80) > 0
                 Flags.Z = Regs8.A = 0
                 Flags.H = False
                 Flags.P = Parity(Regs8.A)
                 Flags.N = False
      Case &H68:                                                 ' in l,(c)
      Case &H69:                                                 ' out (c),l
      Case &H6A: Call Adc16(Regs8.H, Regs8.L, Regs8.H, Regs8.L)  ' adc hl,hl
      Case &H6F: data1 = Read8(Regs8.H, Regs8.L)                 ' rld
                 Data2 = data1 Mod 16   ' data2 holds low nibble of (HL)
                 data1 = data1 \ 16    ' data1 holds high nibble of (HL)
                 Data3 = Regs8.A And &HF ' data3 hold low nibble of A
                 Regs8.A = (Regs8.A And &HF0) + data1 ' high nibble of (HL) in low nibble A
                 data1 = data1 + Data3 * 16 ' data1 now has low of (hl) in high and low of A in low
                 Call Write8(Regs8.H, Regs8.L, data1)
                 Flags.S = (Regs8.A And &H80) > 0
                 Flags.Z = Regs8.A = 0
                 Flags.H = False
                 Flags.P = Parity(Regs8.A)
                 Flags.N = False
      Case &H72: Call Sbc16(Regs8.S, Regs8.P)                    ' sbc hl,sp
      Case &H73: data1 = NextByte                                ' ld (nn),sp
                 Data2 = NextByte
                 Call Write16(data1, Data2, Regs16.SP \ 256, Regs16.SP And &HFF)
      Case &H78:                                                 ' in a,(c)
      Case &H79:                                                 ' out (c),a
      Case &H7A: Call Sbc16(Regs8.S, Regs8.P)                    ' adc hl,sp
      Case &H7B: data1 = NextByte                                ' ld sp,(nn)
                 Data2 = NextByte
                 Regs16.SP = Read16(data1, Data2)
      Case &HA0: Temp1 = (CLng(Regs8.B) * 256 + Regs8.C)         ' ldi
                 Temp2 = (CLng(Regs8.D) * 256 + Regs8.E)
                 Temp3 = (CLng(Regs8.H) * 256 + Regs8.L)
                 data1 = Read8((Temp3 And CLng("&HFF00")) / 256, Temp3 And &HFF)
                 Call Write8((Temp2 And CLng("&HFF00")) / 256, Temp2 And &HFF, data1)
                 Temp1 = Temp1 - 1
                 Temp2 = (Temp2 + 1) And CLng("&HFFFF")
                 Temp3 = (Temp3 + 1) And CLng("&HFFFF")
                 Flags.H = False
                 Flags.P = (Temp1 = 0)
                 Flags.N = False
                 Regs8.B = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.C = Temp1 And &HFF
                 Regs8.D = (Temp2 And CLng("&HFF00")) / 256
                 Regs8.E = Temp2 And &HFF
                 Regs8.H = (Temp3 And CLng("&HFF00")) / 256
                 Regs8.L = Temp3 And &HFF
      Case &HA1: Temp1 = CLng(Regs8.B) * 256 + Regs8.C
                 Temp2 = CLng(Regs8.H) * 256 + Regs8.L
                 Temp1 = (Temp2 - 1) And CLng("&HFFFF")
                 Temp2 = (Temp3 + 1) And CLng("&HFFFF")
                 Call Cp(6)
                 Regs8.H = Temp2 \ 256
                 Regs8.L = Temp2 Mod 256
                 Regs8.B = Temp1 \ 256
                 Regs8.C = Temp1 Mod 256
      Case &HA2:                                                 ' ini
      Case &HA3:                                                 ' oti
      Case &HA8: Temp1 = (CLng(Regs8.B) * 256 + Regs8.C)         ' ldd
                 Temp2 = (CLng(Regs8.D) * 256 + Regs8.E)
                 Temp3 = (CLng(Regs8.H) * 256 + Regs8.L)
                 data1 = Read8((Temp3 And CLng("&HFF00")) / 256, Temp3 And &HFF)
                 Call Write8((Temp2 And CLng("&HFF00")) / 256, Temp2 And &HFF, data1)
                 Temp1 = Temp1 - 1
                 Temp2 = (Temp2 - 1) And CLng("&HFFFF")
                 Temp3 = (Temp3 - 1) And CLng("&HFFFF")
                 Flags.H = False
                 Flags.P = (Temp1 = 0)
                 Flags.N = False
                 Regs8.B = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.C = Temp1 And &HFF
                 Regs8.D = (Temp2 And CLng("&HFF00")) / 256
                 Regs8.E = Temp2 And &HFF
                 Regs8.H = (Temp3 And CLng("&HFF00")) / 256
                 Regs8.L = Temp3 And &HFF
      Case &HA9:                                                 ' cpd
                 Temp1 = CLng(Regs8.B * 256) + Regs8.C  ' temp1 = bc
                 Temp2 = CLng(Regs8.H * 256) + Regs8.L  ' temp2 = hl
                 Temp3 = Regs8.A - RAM(Temp2)          ' temp3 = compare result
                 Temp1 = (Temp1 - 1) And CLng("&hFFFF")                     ' bc = bc - 1
                 Temp2 = (Temp2 - 1) And CLng("&HFFFF")  ' hl = hl - 1
                 Flags.Z = (Temp3 = 0)
                 Flags.P = (Temp1 = 0)
                 Flags.S = (Temp3 And &H80) > 0
                
      Case &HAA:                                                 ' ind
      Case &HAB:                                                 ' outd
      Case &HB0: Temp1 = (CLng(Regs8.B) * 256 + Regs8.C)         ' ldir
                 Temp2 = (CLng(Regs8.D) * 256 + Regs8.E)
                 Temp3 = (CLng(Regs8.H) * 256 + Regs8.L)
                 While Temp1 > 0
                    data1 = Read8((Temp3 And CLng("&HFF00")) / 256, Temp3 And &HFF)
                    Call Write8((Temp2 And CLng("&HFF00")) / 256, Temp2 And &HFF, data1)
                    Temp1 = Temp1 - 1
                    Temp2 = (Temp2 + 1) And CLng("&HFFFF")
                    Temp3 = (Temp3 + 1) And CLng("&HFFFF")
                 Wend
                 Flags.H = False
                 Flags.P = False
                 Flags.N = False
                 Regs8.B = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.C = Temp1 And &HFF
                 Regs8.D = (Temp2 And CLng("&HFF00")) / 256
                 Regs8.E = Temp2 And &HFF
                 Regs8.H = (Temp3 And CLng("&HFF00")) / 256
                 Regs8.L = Temp3 And &HFF
      Case &HB1:                                                 ' cpir
      Case &HB2:  Flags.Z = 1                                    ' inir
                  Flags.N = 1
      Case &HB3:                                                 ' otir
      Case &HB8: Temp1 = (CLng(Regs8.B) * 256 + Regs8.C)         ' lddr
                 Temp2 = (CLng(Regs8.D) * 256 + Regs8.E)
                 Temp3 = (CLng(Regs8.H) * 256 + Regs8.L)
                 While Temp1 > 0
                    data1 = Read8((Temp3 And CLng("&HFF00")) / 256, Temp3 And &HFF)
                    Call Write8((Temp2 And CLng("&HFF00")) / 256, Temp2 And &HFF, data1)
                    Temp1 = Temp1 - 1
                    Temp2 = (Temp2 - 1) And CLng("&HFFFF")
                    Temp3 = (Temp3 - 1) And CLng("&HFFFF")
                 Wend
                 Flags.H = False
                 Flags.P = False
                 Flags.N = False
                 Regs8.B = (Temp1 And CLng("&HFF00")) / 256
                 Regs8.C = Temp1 And &HFF
                 Regs8.D = (Temp2 And CLng("&HFF00")) / 256
                 Regs8.E = Temp2 And &HFF
                 Regs8.H = (Temp3 And CLng("&HFF00")) / 256
                 Regs8.L = Temp3 And &HFF
      Case &HB9:                                             ' cpdr
                 Temp1 = CLng(Regs8.B * 256) + Regs8.C  ' temp1 = bc
                 Temp2 = 1
                 Temp3 = 1
                 While (Temp2 > 0) And (Temp3 <> 0)
                    Temp2 = CLng(Regs8.H * 256) + Regs8.L  ' temp2 = hl
                    Temp3 = Regs8.A - RAM(Temp2)          ' temp3 = compare result
                    Temp1 = Temp1 - 1                     ' bc = bc - 1
                    Temp2 = (Temp2 - 1) And CLng(&HFFFF)  ' hl = hl - 1
                 Wend
                 Flags.Z = (Temp3 = 0)
                 Flags.P = (Temp1 = 0)
                 Flags.S = (Temp3 And &H80) > 0
      Case &HBA:                                                            ' indr
      Case &HBB:                                                          ' otdr
      Case Else  ' bad opcode
   End Select
End Sub


