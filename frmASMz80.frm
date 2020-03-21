VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmASMz80 
   Caption         =   "ASMz80"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optBinSel 
      Caption         =   "BINary output"
      Height          =   510
      Left            =   9210
      TabIndex        =   9
      Top             =   1700
      Width           =   1905
   End
   Begin VB.OptionButton optCPMcom 
      Caption         =   "CPM Executable"
      Height          =   510
      Left            =   9225
      TabIndex        =   7
      Top             =   1300
      Width           =   1905
   End
   Begin VB.OptionButton optIntelHex 
      Caption         =   "Intel Hex Output"
      Height          =   510
      Left            =   9210
      TabIndex        =   6
      Top             =   900
      Width           =   1905
   End
   Begin VB.CheckBox chkListing 
      Caption         =   "Enable Assembly Listing"
      Height          =   495
      Left            =   45
      TabIndex        =   5
      Top             =   1620
      Width           =   2190
   End
   Begin VB.CommandButton cmdAssemble 
      Caption         =   "Assemble"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      TabIndex        =   4
      Top             =   585
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtObject 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2355
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1365
      Width           =   6780
   End
   Begin VB.TextBox txtList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2355
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   705
      Width           =   6800
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2355
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   105
      Width           =   6800
   End
   Begin VB.CommandButton cmdSelectFile 
      Caption         =   "Select File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1695
   End
   Begin VB.Label Label1 
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   1170
      Width           =   2355
   End
End
Attribute VB_Name = "frmASMz80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAssemble_Click()
Dim SourceLine As String
Dim OBJStr As String
Dim OBJStr1 As String
Dim tempstr As String
Dim ct As Long
  
   
   Label1.Caption = "Starting Assembly"
   frmAsmList.txtAsmList.Text = ""
   Call InitVars
   If optCPMcom Then
      ASMPC = &H100
   Else
      ASMPC = 0
   End If
   Pass2 = False
   EndOp = False
   SrcNum = FreeFile
   Open txtSource.Text For Input As #SrcNum
   While Not EOF(1)
      Line Input #SrcNum, SourceLine
      OBJStr = Assemble(SourceLine)
   Wend
   Close #1
   If optCPMcom Then
      ASMPC = &H100
   Else
      ASMPC = 0
   End If
   Pass2 = True
   ASMPC = 0
   SrcNum = FreeFile
   Open txtSource.Text For Input As #SrcNum
   ListNum = FreeFile
   Open txtList.Text For Output As #ListNum
   ObjectNum = FreeFile
   If HexSel Then
      Open txtObject.Text For Output As #ObjectNum
   ElseIf CPMSel Then
      Open txtObject.Text For Binary As #ObjectNum
   ElseIf BINSel Then
      Open txtObject.Text For Binary As #ObjectNum
   End If
   
   Call InitObject
   If chkListing.value = 1 Then
      frmAsmList.Show
   End If
   
   Print #ListNum, txtSource.Text
   While Not EOF(SrcNum) And (Not EndOp)
      Line Input #SrcNum, SourceLine
      OBJStr = Assemble(SourceLine)
      OBJStr1 = OBJStr
      If Len(OBJStr) > 5 Then Call ObjectSave(OBJStr)
      Print #ListNum, Left(OBJStr, 4) & "  " & Left(Right(OBJStr, Len(OBJStr) - 4) + "          ", 8) & "  " & SourceLine + vbCrLf;
      If chkListing Then frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Left(OBJStr, 4) + " " + Left(Right(OBJStr, Len(OBJStr) - 4) + "          ", 8) + " " + SourceLine + vbCrLf
      If Len(OBJStr) > 16 Then
         Mid(OBJStr, 1, 12) = String(12, " ")
         OBJStr = Trim(OBJStr)
         While Len(OBJStr) > 0
            If Len(OBJStr) > 0 Then
               Print #ListNum, "      " + Left(OBJStr, 8)
               If chkListing Then frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + "     " + Left(OBJStr + "          ", 8) + vbCrLf
            End If
            Mid(OBJStr, 1, 8) = "        "
            OBJStr = Trim(OBJStr)
         Wend
      End If
      If ErrorStr <> "" Then
         Print #ListNum, ErrorStr
         If chkListing Then frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + ErrorStr + vbCrLf
      End If
      DoEvents
      ErrorStr = ""
   Wend

   Call ObjectClose(OBJStr1)
   If chkListing Then frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + vbCrLf + vbCrLf + "Symbol Table" + vbCrLf
   Print #ListNum,
   Print #ListNum, "Symbol Table"
   For ct = 1 To SymTabLast - 1
      If chkListing Then
         frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Left(symTable(ct).name + "                  ", 16)
         frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Right("0000" + Hex(symTable(ct).value), 4) + "  "
         If Not symTable(ct).defined Then frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + "Undefined  "
         If symTable(ct).multiDef Then frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + "Multple Defined  "
         frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + vbCrLf
      End If
      Print #ListNum, Left(symTable(ct).name + "                  ", 16);
      Print #ListNum, Right("0000" + Hex(symTable(ct).value), 4) + "  ";
      If Not symTable(ct).defined Then Print #ListNum, "Undefined  ";
      If symTable(ct).multiDef Then Print #ListNum, "Multple Defined  ";
      Print #ListNum, " "
      
   Next ct

   Close SrcNum
   Close ListNum
   Close ObjectNum
   Label1.Caption = "Assembly Complete"
   cmdAssemble.Visible = True

End Sub

Private Sub cmdSelectFile_Click()
   
   CommonDialog1.Filter = "z80 Assembly File|*.z80|Assembly File|*.asm|All FIles|*.*"
   CommonDialog1.ShowOpen
   
   If Len(CommonDialog1.FileName) > 0 Then
      HexSel = False
      CPMSel = False
      BINSel = False
      
      txtSource.Text = CommonDialog1.FileName
      txtList.Text = Left(txtSource.Text, Len(txtSource.Text) - 4) + ".LST"
      If optIntelHex Then
         txtObject.Text = Left(txtSource.Text, Len(txtSource.Text) - 4) + ".HEX"
         HexSel = True
      ElseIf optCPMcom Then
         txtObject.Text = Left(txtSource.Text, Len(txtSource.Text) - 4) + ".com"
         CPMSel = True
      ElseIf optCPMcom Then
         txtObject.Text = Left(txtSource.Text, Len(txtSource.Text) - 4) + ".BIN"
         BINSel = True
      End If
      cmdAssemble.Visible = True
   End If
   frmAsmList.txtAsmList.Text = ""
End Sub

Private Sub Form_Load()
Dim ct As Long
Dim TestLine As String

   Load frmAsmList

   cmdAssemble.Visible = False
   frmASMz80.Height = 2600
   optIntelHex.value = True
   txtSource.Text = ""
   txtList.Text = ""
   txtObject.Text = ""
   
   IntelHexAddr = -1
   IntelHexInit = False
   CPMInit = False
   BinInit = False

End Sub

Private Sub Form_Resize()
   If frmASMz80.Height < 2600 Then
      frmASMz80.Height = 2600
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload frmAsmList
   Close
End Sub

Private Sub optCPMcom_Click()
   If txtSource.Text <> "" Then
      txtObject.Text = Left(txtSource.Text, Len(txtSource.Text) - 4) + ".com"
      HexSel = False
      CPMSel = True
      BINSel = False
   End If

End Sub

Private Sub optIntelHex_Click()
   If txtSource.Text <> "" Then
      txtObject.Text = Left(txtSource.Text, Len(txtSource.Text) - 4) + ".HEX"
      HexSel = True
      CPMSel = False
      BINSel = False
   End If

End Sub

Private Sub optBinSel_Click()
   If txtSource.Text <> "" Then
      txtObject.Text = Left(txtSource.Text, Len(txtSource.Text) - 4) + ".BIN"
      HexSel = False
      CPMSel = False
      BINSel = True
   End If

End Sub

Private Sub chkListing_Click()
   If chkListing.value = 1 Then
   End If
End Sub

