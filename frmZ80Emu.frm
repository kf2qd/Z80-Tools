VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmz80Emu 
   Caption         =   "z80 Emulator"
   ClientHeight    =   11055
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   19815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   19815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   5580
      MaxLength       =   4
      TabIndex        =   92
      Text            =   "0000"
      Top             =   540
      Width           =   870
   End
   Begin VB.TextBox txtLRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   4635
      MaxLength       =   4
      TabIndex        =   90
      Text            =   "0000"
      Top             =   540
      Width           =   870
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   17
      Left            =   2955
      MaxLength       =   2
      TabIndex        =   88
      Text            =   "00"
      Top             =   2190
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   16
      Left            =   2475
      MaxLength       =   2
      TabIndex        =   87
      Text            =   "00"
      Top             =   2190
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   14
      Left            =   3555
      MaxLength       =   2
      TabIndex        =   85
      Text            =   "00"
      Top             =   1350
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   15
      Left            =   4035
      MaxLength       =   2
      TabIndex        =   84
      Text            =   "00"
      Top             =   1350
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   3555
      MaxLength       =   2
      TabIndex        =   82
      Text            =   "00"
      Top             =   540
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   7
      Left            =   4065
      MaxLength       =   2
      TabIndex        =   81
      Text            =   "00"
      Top             =   540
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   12
      Left            =   2475
      MaxLength       =   2
      TabIndex        =   79
      Text            =   "00"
      Top             =   1350
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   13
      Left            =   2955
      MaxLength       =   2
      TabIndex        =   78
      Text            =   "00"
      Top             =   1350
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   2475
      MaxLength       =   2
      TabIndex        =   76
      Text            =   "00"
      Top             =   540
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   5
      Left            =   2985
      MaxLength       =   2
      TabIndex        =   75
      Text            =   "00"
      Top             =   540
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   8
      Left            =   360
      MaxLength       =   2
      TabIndex        =   72
      Text            =   "00"
      Top             =   1335
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   9
      Left            =   840
      MaxLength       =   2
      TabIndex        =   71
      Text            =   "00"
      Top             =   1335
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   10
      Left            =   1395
      MaxLength       =   2
      TabIndex        =   70
      Text            =   "00"
      Top             =   1335
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   11
      Left            =   1875
      MaxLength       =   2
      TabIndex        =   69
      Text            =   "00"
      Top             =   1335
      Width           =   450
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Command1"
      Height          =   510
      Left            =   4680
      TabIndex        =   68
      Top             =   1260
      Width           =   1680
   End
   Begin VB.TextBox txtLRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   1425
      MaxLength       =   4
      TabIndex        =   67
      Text            =   "0000"
      Top             =   2195
      Width           =   870
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   66
      Text            =   "00"
      Top             =   540
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   1395
      MaxLength       =   2
      TabIndex        =   65
      Text            =   "00"
      Top             =   520
      Width           =   450
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   900
      MaxLength       =   2
      TabIndex        =   64
      Text            =   "00"
      Top             =   540
      Width           =   450
   End
   Begin VB.CommandButton cmdClearMem 
      Caption         =   "Clear Memory"
      Height          =   495
      Left            =   15555
      TabIndex        =   62
      Top             =   6810
      Width           =   1635
   End
   Begin VB.CheckBox chkShowAssembly 
      Caption         =   "Show Assembly"
      Height          =   495
      Left            =   8300
      TabIndex        =   60
      Top             =   1650
      Width           =   1215
   End
   Begin VB.TextBox txtNumSteps 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   15885
      MaxLength       =   4
      TabIndex        =   59
      Text            =   "5"
      Top             =   4275
      Width           =   990
   End
   Begin VB.CommandButton cmdMultiStep 
      Caption         =   "Multi-Step"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15885
      TabIndex        =   58
      Top             =   3870
      Width           =   1875
   End
   Begin VB.CommandButton cmdSaveSource 
      Caption         =   "Save Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   19620
      TabIndex        =   57
      Top             =   1530
      Width           =   2000
   End
   Begin VB.TextBox txtAsmList 
      Height          =   465
      Left            =   9450
      TabIndex        =   56
      Top             =   585
      Width           =   4000
   End
   Begin VB.CommandButton cmdAssembleFile 
      Caption         =   "Assemble File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6525
      TabIndex        =   55
      Top             =   990
      Width           =   1590
   End
   Begin VB.CommandButton cmdSelSOurce 
      Caption         =   "Source File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   6525
      TabIndex        =   54
      Top             =   180
      Width           =   1590
   End
   Begin VB.TextBox txtInlineAsmAddr 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   15570
      MaxLength       =   4
      TabIndex        =   53
      Text            =   "0000"
      Top             =   6240
      Width           =   990
   End
   Begin VB.TextBox txtAsmImmed 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   15570
      TabIndex        =   52
      Top             =   5700
      Width           =   4000
   End
   Begin VB.Frame FrmSaveAs 
      Caption         =   "Save As"
      Height          =   1500
      Left            =   13545
      TabIndex        =   47
      Top             =   90
      Width           =   1560
      Begin VB.OptionButton optBINbin 
         Caption         =   "BIN"
         Height          =   465
         Left            =   90
         TabIndex        =   51
         Top             =   585
         Width           =   1230
      End
      Begin VB.OptionButton optCPMcom 
         Caption         =   "CPM"
         Height          =   465
         Left            =   90
         TabIndex        =   50
         Top             =   225
         Width           =   1230
      End
      Begin VB.OptionButton optIntelHex 
         Caption         =   "HEX"
         Height          =   465
         Left            =   90
         TabIndex        =   49
         Top             =   990
         Width           =   1230
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   465
         Index           =   1
         Left            =   90
         TabIndex        =   48
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.TextBox txtAsmObj 
      Height          =   465
      Left            =   9450
      TabIndex        =   46
      Top             =   1080
      Width           =   4000
   End
   Begin VB.TextBox txtAsmSource 
      Height          =   465
      Left            =   9450
      TabIndex        =   45
      Top             =   90
      Width           =   4000
   End
   Begin VB.TextBox txtCtrlAdd 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   36
      Text            =   "0000"
      Top             =   7560
      Width           =   1000
   End
   Begin VB.TextBox txtCtrlWord 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6645
      MaxLength       =   1
      TabIndex        =   35
      Text            =   "I"
      Top             =   7560
      Width           =   480
   End
   Begin VB.TextBox txtLabelAdd 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9615
      MaxLength       =   4
      TabIndex        =   34
      Text            =   "0000"
      Top             =   7515
      Width           =   1020
   End
   Begin VB.TextBox txtLabelName 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10845
      TabIndex        =   33
      Text            =   " "
      Top             =   7530
      Width           =   2250
   End
   Begin VB.TextBox txtSaveEnd 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   17550
      MaxLength       =   4
      TabIndex        =   32
      Top             =   2745
      Width           =   1020
   End
   Begin VB.TextBox txtSaveStart 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   17550
      MaxLength       =   4
      TabIndex        =   31
      Top             =   2340
      Width           =   1020
   End
   Begin VB.CommandButton cmdSaveHex 
      Caption         =   "Save Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   17775
      TabIndex        =   28
      Top             =   1530
      Width           =   2000
   End
   Begin VB.CommandButton cmdSaveBin 
      Caption         =   "Save Bin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   17775
      TabIndex        =   27
      Top             =   1050
      Width           =   2000
   End
   Begin VB.CommandButton cmdSaveCPM 
      Caption         =   "Save CPM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   17730
      TabIndex        =   26
      Top             =   570
      Width           =   2000
   End
   Begin VB.TextBox txtLRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   360
      MaxLength       =   4
      TabIndex        =   25
      Text            =   "0000"
      Top             =   2195
      Width           =   870
   End
   Begin VB.CommandButton cmdLoadHEX 
      Caption         =   "Load Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   15975
      TabIndex        =   24
      Top             =   1530
      Width           =   1700
   End
   Begin VB.CommandButton cmdLoadCPM 
      Caption         =   "Load CPM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15975
      TabIndex        =   23
      Top             =   570
      Width           =   1700
   End
   Begin VB.TextBox txtStart 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6735
      MaxLength       =   4
      TabIndex        =   9
      Text            =   "0000"
      Top             =   2115
      Width           =   750
   End
   Begin VB.TextBox txtHexdisp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Left            =   5220
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmZ80Emu.frx":0000
      Top             =   2565
      Width           =   9990
   End
   Begin VB.CommandButton cmdSingleStep 
      Caption         =   "Single Step"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15885
      TabIndex        =   1
      Top             =   3375
      Width           =   1875
   End
   Begin VB.CommandButton cmdLoadBin 
      Caption         =   "Load Bin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15975
      TabIndex        =   0
      Top             =   1050
      Width           =   1700
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4620
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDissasm 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   330
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   10
      Top             =   2835
      Width           =   4755
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   390
      MaxLength       =   2
      TabIndex        =   63
      Text            =   "00"
      Top             =   520
      Width           =   450
   End
   Begin VB.Line Line11 
      BorderWidth     =   4
      X1              =   6480
      X2              =   6480
      Y1              =   210
      Y2              =   1035
   End
   Begin VB.Line Line10 
      BorderWidth     =   4
      X1              =   5535
      X2              =   5535
      Y1              =   210
      Y2              =   1035
   End
   Begin VB.Label Label14 
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5850
      TabIndex        =   93
      Top             =   270
      Width           =   330
   End
   Begin VB.Label Label13 
      Caption         =   "IX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4875
      TabIndex        =   91
      Top             =   270
      Width           =   330
   End
   Begin VB.Label Label11 
      Caption         =   "I       R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2625
      TabIndex        =   89
      Top             =   1935
      Width           =   750
   End
   Begin VB.Label Label9 
      Caption         =   "H'    L'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3705
      TabIndex        =   86
      Top             =   1095
      Width           =   750
   End
   Begin VB.Label Label8 
      Caption         =   "H    L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3735
      TabIndex        =   83
      Top             =   285
      Width           =   750
   End
   Begin VB.Line Line9 
      BorderWidth     =   4
      X1              =   4575
      X2              =   4575
      Y1              =   225
      Y2              =   1860
   End
   Begin VB.Label Label7 
      Caption         =   "D'    E'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2625
      TabIndex        =   80
      Top             =   1095
      Width           =   750
   End
   Begin VB.Label Label6 
      Caption         =   "D    E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2655
      TabIndex        =   77
      Top             =   285
      Width           =   750
   End
   Begin VB.Line Line7 
      BorderWidth     =   4
      X1              =   3495
      X2              =   3495
      Y1              =   225
      Y2              =   2685
   End
   Begin VB.Label Label5 
      Caption         =   "A'    F'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   510
      TabIndex        =   74
      Top             =   1080
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "B'   C'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1545
      TabIndex        =   73
      Top             =   1080
      Width           =   750
   End
   Begin VB.Label lblAsmStatus 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10215
      TabIndex        =   61
      Top             =   1620
      Width           =   2655
   End
   Begin VB.Label lblImmedAsm 
      Caption         =   "Immediate Assembly"
      Height          =   435
      Left            =   15585
      TabIndex        =   44
      Top             =   5175
      Width           =   1005
   End
   Begin VB.Label lblObjFile 
      Caption         =   "OBJ File"
      Height          =   345
      Left            =   8190
      TabIndex        =   43
      Top             =   1125
      Width           =   1230
   End
   Begin VB.Label lblListFile 
      Caption         =   "List File"
      Height          =   345
      Left            =   8190
      TabIndex        =   42
      Top             =   630
      Width           =   1230
   End
   Begin VB.Label lblAsmSource 
      Caption         =   "ASM Source"
      Height          =   345
      Left            =   8190
      TabIndex        =   41
      Top             =   135
      Width           =   1230
   End
   Begin VB.Label Label31 
      Caption         =   "Control Address"
      Height          =   495
      Left            =   5280
      TabIndex        =   40
      Top             =   7110
      Width           =   1215
   End
   Begin VB.Label Label30 
      Caption         =   "A - ASCII       B - Byte        E - End         I - Instruction  S - Storage     W - Word        X - Clear"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   7485
      TabIndex        =   39
      Top             =   7335
      Width           =   1935
   End
   Begin VB.Label Label29 
      Caption         =   "Symbol Address"
      Height          =   495
      Left            =   9615
      TabIndex        =   38
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Symbol Name"
      Height          =   495
      Left            =   10905
      TabIndex        =   37
      Top             =   7065
      Width           =   900
   End
   Begin VB.Label lblEndAddr 
      Caption         =   "End Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15960
      TabIndex        =   30
      Top             =   2775
      Width           =   1485
   End
   Begin VB.Label lblStartAddr 
      Caption         =   "Start Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15960
      TabIndex        =   29
      Top             =   2400
      Width           =   1485
   End
   Begin VB.Line Line5 
      BorderWidth     =   4
      X1              =   1360
      X2              =   1360
      Y1              =   225
      Y2              =   2700
   End
   Begin VB.Line Line4 
      BorderWidth     =   4
      X1              =   300
      X2              =   3495
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label lblCflg 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4785
      TabIndex        =   22
      Top             =   2295
      Width           =   150
   End
   Begin VB.Label lblNflg 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   21
      Top             =   2295
      Width           =   150
   End
   Begin VB.Label lblPflg 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4290
      TabIndex        =   20
      Top             =   2310
      Width           =   150
   End
   Begin VB.Label lblHflg 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4065
      TabIndex        =   19
      Top             =   2310
      Width           =   150
   End
   Begin VB.Label lblZflg 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   18
      Top             =   2310
      Width           =   150
   End
   Begin VB.Label lblSflg 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3615
      TabIndex        =   17
      Top             =   2310
      Width           =   150
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4785
      TabIndex        =   16
      Top             =   2085
      Width           =   150
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   15
      Top             =   2085
      Width           =   150
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4290
      TabIndex        =   14
      Top             =   2100
      Width           =   150
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4065
      TabIndex        =   13
      Top             =   2100
      Width           =   150
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   12
      Top             =   2100
      Width           =   150
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3615
      TabIndex        =   11
      Top             =   2100
      Width           =   150
   End
   Begin VB.Label Label2 
      Caption         =   "Start Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5280
      TabIndex        =   8
      Top             =   2175
      Width           =   1485
   End
   Begin VB.Line Line8 
      BorderWidth     =   4
      X1              =   2430
      X2              =   2430
      Y1              =   210
      Y2              =   2685
   End
   Begin VB.Line Line6 
      BorderWidth     =   4
      X1              =   320
      X2              =   320
      Y1              =   200
      Y2              =   2715
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   300
      X2              =   4575
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   315
      X2              =   6480
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   315
      X2              =   6480
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Label lblRegA 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   420
      TabIndex        =   6
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label18 
      Caption         =   "PC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   615
      TabIndex        =   5
      Top             =   1955
      Width           =   435
   End
   Begin VB.Label Label16 
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1710
      TabIndex        =   4
      Top             =   1955
      Width           =   330
   End
   Begin VB.Label Label3 
      Caption         =   "B    C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Top             =   270
      Width           =   750
   End
   Begin VB.Label lblA 
      Caption         =   "A     F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   540
      TabIndex        =   2
      Top             =   270
      Width           =   750
   End
End
Attribute VB_Name = "frmz80Emu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DisplayToggle As Integer
Public Srcfile As Long

Private Sub Form_Load()
Dim X, Y, Z As Long

   For X = 0 To CLng("&hFFFF")
      RAM(X) = 0
      Memory(X).Label = ""
      Memory(X).Usage = 0
   Next X

   Label30.Caption = "A - ASCII" + vbCrLf + "B - Byte" + vbCrLf + "E - END" + vbCrLf + "I - Instruction" + vbCrLf + "S - Storage" + vbCrLf + "W - Word" + vbCrLf + "X - Clear"
   Call DisplayRegs
   txtHexdisp.Text = ""

   txtStart.Text = "0000"
   
   EndAddress = CLng(&H10000)
   
   frmz80Emu.Height = 9500
   frmz80Emu.Width = 15500
   ' emulator command buttons
   cmdSingleStep.Visible = True
   cmdSingleStep.Enabled = True
   cmdSingleStep.Left = 6495
   cmdSingleStep.Top = 150
   cmdMultiStep.Visible = True
   cmdMultiStep.Enabled = True
   cmdMultiStep.Left = 6495
   cmdMultiStep.Top = 650
   txtNumSteps.Visible = True
   txtNumSteps.Top = 650
   txtNumSteps.Left = 8500
   cmdClearMem.Visible = True
   cmdClearMem.Left = 13400
   cmdClearMem.Top = 150
   cmdLoadCPM.Visible = False
   cmdLoadCPM.Enabled = False
   cmdLoadCPM.Left = 8000
   cmdLoadCPM.Top = 125
   cmdLoadBin.Visible = False
   cmdLoadBin.Enabled = False
   cmdLoadBin.Left = 8000
   cmdLoadBin.Top = 600
   cmdLoadHEX.Visible = False
   cmdLoadHEX.Enabled = False
   cmdLoadHEX.Left = 8000
   cmdLoadHEX.Top = 1075
   cmdSaveCPM.Visible = False
   cmdSaveCPM.Enabled = False
   cmdSaveCPM.Left = 9800
   cmdSaveCPM.Top = 125
   cmdSaveBin.Visible = False
   cmdSaveBin.Enabled = False
   cmdSaveBin.Left = 9800
   cmdSaveBin.Top = 600
   cmdSaveHex.Visible = False
   cmdSaveHex.Enabled = False
   cmdSaveHex.Left = 9800
   cmdSaveHex.Top = 1075
   cmdSaveSource.Visible = False
   cmdSaveSource.Enabled = False
   cmdSaveSource.Left = 9800
   cmdSaveSource.Top = 1550
   lblStartAddr.Visible = False
   lblStartAddr.Left = 12000
   lblStartAddr.Top = 285
   lblEndAddr.Visible = False
   lblEndAddr.Left = 12000
   lblEndAddr.Top = 660
   txtSaveStart.Visible = False
   txtSaveStart.Left = 13500
   txtSaveStart.Top = 225
   txtSaveEnd.Visible = False
   txtSaveEnd.Left = 13500
   txtSaveEnd.Top = 630
   lblImmedAsm.Visible = True
   lblImmedAsm.Enabled = True
   lblImmedAsm.Left = 8190
   lblImmedAsm.Top = 2000
   txtInlineAsmAddr.Visible = True
   txtInlineAsmAddr.Enabled = True
   txtInlineAsmAddr.Text = "0000"
   txtInlineAsmAddr.Left = 9200
   txtInlineAsmAddr.Top = 2000
   txtAsmImmed.Visible = True
   txtAsmImmed.Left = 10300
   txtAsmImmed.Top = 2000
   chkShowAssembly.Visible = False
   chkShowAssembly.Top = 1650
   chkShowAssembly.Left = 8200
   DisplayToggle = 1
   cmdSwitch.Caption = "Show Load/Save"
   
   ' assembler command buttons
   cmdSelSOurce.Visible = False
   cmdSelSOurce.Left = 6480
   cmdSelSOurce.Top = 180
   cmdAssembleFile.Visible = False
   cmdAssembleFile.Left = 6480
   cmdAssembleFile.Top = 990
   lblAsmSource.Visible = False
   lblAsmSource.Enabled = False
   lblAsmSource.Left = 8190
   lblAsmSource.Top = 135
   chkShowAssembly.Visible = False
   chkShowAssembly.Top = 1500
   chkShowAssembly.Left = 8200
   lblAsmStatus.Visible = False
   lblAsmStatus.Left = 10215
   lblAsmStatus.Top = 1620
   lblAsmStatus.Caption = ""

   lblListFile.Visible = False
   lblListFile.Enabled = False
   lblListFile.Left = 8190
   lblListFile.Top = 630
   lblObjFile.Visible = False
   lblObjFile.Enabled = False
   lblObjFile.Left = 8190
   lblObjFile.Top = 1125
   txtAsmSource.Visible = False
   txtAsmSource.Left = 9450
   txtAsmSource.Top = 90
   txtAsmList.Visible = False
   txtAsmList.Left = 9450
   txtAsmList.Top = 585
   txtAsmObj.Visible = False
   txtAsmObj.Left = 9450
   txtAsmObj.Top = 1080
   FrmSaveAs.Visible = False
   optIntelHex.value = True
   FrmSaveAs.Left = 13545
   FrmSaveAs.Top = 45
   
   frmAsmList.txtAsmList.Text = ""
   
   Call DisplayHex
   Call DisplayAssembly
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload frmAsmList
   Unload frmConsoleDisplay
End Sub

Private Sub cmdSwitch_Click()

   DisplayToggle = (DisplayToggle + 1)
   If DisplayToggle = 3 Then DisplayToggle = 0
   
    ' option '0' Assembler selection
   If DisplayToggle = 0 Then
      cmdSwitch.Caption = "Show Execute"
      cmdSingleStep.Visible = False
      cmdMultiStep.Visible = False
      txtNumSteps.Visible = False
      cmdClearMem.Visible = False
      cmdLoadCPM.Visible = False
      cmdLoadCPM.Enabled = False
      cmdLoadBin.Visible = False
      cmdLoadBin.Enabled = False
      cmdLoadHEX.Visible = False
      cmdLoadHEX.Enabled = False
      cmdSaveCPM.Visible = False
      cmdSaveCPM.Enabled = False
      cmdSaveBin.Visible = False
      cmdSaveBin.Enabled = False
      cmdSaveHex.Visible = False
      cmdSaveHex.Enabled = False
      cmdSaveSource.Visible = False
      cmdSaveSource.Enabled = False
      lblStartAddr.Visible = False
      lblEndAddr.Visible = False
      txtSaveStart.Visible = False
      txtSaveEnd.Visible = False
      lblImmedAsm.Visible = False
      lblImmedAsm.Enabled = False
      txtAsmImmed.Visible = False
      txtInlineAsmAddr.Visible = False
      txtInlineAsmAddr.Enabled = False
      
      cmdSelSOurce.Visible = True
      cmdAssembleFile.Visible = True
      lblAsmSource.Visible = True
      lblAsmSource.Enabled = True
      lblListFile.Visible = True
      lblListFile.Enabled = True
      lblObjFile.Visible = True
      lblObjFile.Enabled = True
      chkShowAssembly.Visible = True
      txtAsmSource.Visible = True
      txtAsmList.Visible = True
      txtAsmObj.Visible = True
      FrmSaveAs.Visible = True
      lblAsmStatus.Visible = True
      
      ' option '1' Execute
   ElseIf DisplayToggle = 1 Then
   
      cmdSwitch.Caption = "Show Load/Save"
      
      cmdSelSOurce.Visible = False
      cmdAssembleFile.Visible = False
      cmdSingleStep.Visible = True
      cmdMultiStep.Visible = True
      txtNumSteps.Visible = True
      cmdClearMem.Visible = True
      lblAsmSource.Visible = False
      lblAsmSource.Enabled = False
      lblListFile.Visible = False
      lblListFile.Enabled = False
      lblObjFile.Visible = False
      lblObjFile.Enabled = False
      txtAsmSource.Visible = False
      txtAsmList.Visible = False
      txtAsmObj.Visible = False
      chkShowAssembly.Visible = True
      lblAsmStatus.Visible = False
      FrmSaveAs.Visible = False
      cmdSingleStep.Visible = True
      cmdLoadCPM.Visible = False
      cmdLoadCPM.Enabled = False
      cmdLoadBin.Visible = False
      cmdLoadBin.Enabled = False
      cmdLoadHEX.Visible = False
      cmdLoadHEX.Enabled = False
      cmdSaveCPM.Visible = False
      cmdSaveBin.Visible = False
      cmdSaveHex.Visible = False
      cmdSaveSource.Visible = False
      lblStartAddr.Visible = False
      lblEndAddr.Visible = False
      txtSaveStart.Visible = False
      txtSaveEnd.Visible = False
      txtAsmImmed.Visible = True
      lblImmedAsm.Visible = True
      lblImmedAsm.Enabled = True
      txtInlineAsmAddr.Visible = True
      txtInlineAsmAddr.Enabled = True
   
      ' option '2' load and save binarys and dissambled code
   ElseIf DisplayToggle = 2 Then
   
      cmdSwitch.Caption = "Show Assembler"
      
      cmdSelSOurce.Visible = False
      cmdAssembleFile.Visible = False
      cmdSingleStep.Visible = False
      cmdMultiStep.Visible = False
      txtNumSteps.Visible = False
      cmdClearMem.Visible = False
      lblAsmSource.Visible = False
      lblAsmSource.Enabled = False
      lblListFile.Visible = False
      lblListFile.Enabled = False
      lblObjFile.Visible = False
      lblObjFile.Enabled = False
      txtAsmSource.Visible = False
      txtAsmList.Visible = False
      txtAsmObj.Visible = False
      FrmSaveAs.Visible = False
      lblAsmStatus.Visible = False

      cmdSingleStep.Visible = False
      cmdLoadCPM.Visible = True
      cmdLoadCPM.Enabled = True
      cmdLoadBin.Visible = True
      cmdLoadBin.Enabled = True
      cmdLoadHEX.Visible = True
      cmdLoadHEX.Enabled = True
      cmdSaveCPM.Visible = True
      cmdSaveBin.Visible = True
      cmdSaveHex.Visible = True
      cmdSaveSource.Visible = True
      
      lblStartAddr.Visible = True
      lblEndAddr.Visible = True
      txtSaveStart.Visible = True
      txtSaveEnd.Visible = True
      txtAsmImmed.Visible = False
      lblImmedAsm.Visible = False
      lblImmedAsm.Enabled = False
      txtInlineAsmAddr.Visible = False
      txtInlineAsmAddr.Enabled = False
      If (txtSaveStart.Text = "") And (txtSaveEnd.Text = "") Then
         cmdSaveCPM.Enabled = False
         cmdSaveBin.Enabled = False
         cmdSaveHex.Enabled = False
      Else
         cmdSaveCPM.Enabled = True
         cmdSaveBin.Enabled = True
         cmdSaveHex.Enabled = True
      End If
   End If

End Sub

Private Sub cmdClearMem_Click()
Dim ct As Long
   For ct = 0 To CLng(65535)
      Memory(ct).Usage = 0
      Memory(ct).Label = ""
      RAM(ct) = 0
   Next ct
   Call DisplayRegs
   Call DisplayHex
   Call DisplayAssembly
End Sub

Private Sub CmdSingleStep_Click()
   Call Emulate
   Call DisplayRegs
   Call DisplayHex
   Call DisplayAssembly
   txtInlineAsmAddr.Text = txtLRegs(0).Text
End Sub

Private Sub cmdMultiStep_Click()
Dim NumSteps As Long
Dim ct As Long
   Halted = False
   NumSteps = CLng(txtNumSteps.Text)
   ct = 0
   While (ct < NumSteps) And (Not Halted)
      Call Emulate
      ct = ct + 1
   Wend
   Call DisplayRegs
   Call DisplayHex
   Call DisplayAssembly
   txtInlineAsmAddr.Text = txtLRegs(0).Text

End Sub

Private Sub cmdSaveHex_Click()
Dim FileNo As Long
Dim startAddr As Long
Dim endAddr As Long
Dim SaveCt As Long
   
   HexSel = True
   Call InitObject
   CommonDialog1.Filter = "Hex File|*.HEX|All FIles|*.*"
   CommonDialog1.ShowSave
   If Len(CommonDialog1.FileName) > 0 Then
      ObjectNum = FreeFile
      Open CommonDialog1.FileName For Output As #ObjectNum
      startAddr = CLng("&H" + txtSaveStart.Text)
      endAddr = CLng("&H" + txtSaveEnd.Text)
      Call InitObject
      For SaveCt = startAddr To endAddr
         Call ObjectSave(Hex4(SaveCt) + hex2L(CLng(RAM(SaveCt))))
      Next SaveCt
      Call ObjectClose("0000-")
   End If
End Sub

' Bin File = each line has a 4 character HEX address
' a blank space and a 2 character HEX Byte Value
' addresses do not need to be consecutive
' 0000 12

Private Sub cmdLoadBin_Click()
Dim fname As String
Dim BinStrt As Long
Dim BinFin As Long
Dim Binctr As Long
Dim bindat As Byte

   
   CommonDialog1.Filter = ".bin files|*.bin|*.*|All files"
   CommonDialog1.ShowOpen
   fname = CommonDialog1.FileName
   Open fname For Binary As #1
   
   Get #1, , BinStrt
   Get #1, , BinFin
   
   For Binctr = BinStrt To BinFin
      Get #1, , bindat
      RAM(Binctr) = bindat
   Next Binctr
   
   Close 1
   txtStart.Text = Hex(BinStrt And CLng(CLng("&HFF00")))
   Call DisplayHex
   Call DisplayAssembly
  
SkipIt:
   On Error GoTo 0
End Sub
 
Private Sub cmdSaveBin_Click()
Dim bindat As Byte
Dim ByteCt As Long
Dim BinStrt As Long
Dim BinFin As Long
   CommonDialog1.Filter = ".bin Files|*.bin||*.*|All files"
   CommonDialog1.ShowSave
   Open CommonDialog1.FileName For Binary As #1
   BinStrt = CLng("&H" + txtSaveStart.Text)
   BinFin = CLng("&H" + txtSaveEnd.Text)
   Put #1, , BinStrt
   Put #1, , BinFin
   For ByteCt = CLng("&H" + txtSaveStart.Text) To CLng("&H" + txtSaveEnd.Text)
      bindat = RAM(ByteCt)
      Put #1, , bindat
   Next ByteCt
   Close 1

End Sub

' Load a CPM .com file and load it starting at address 0100H
' a CPM executable loads at address 0100H and starts execution there also
' file is just 8 bit values to be loaded into memory.
Private Sub cmdLoadCPM_Click()
Dim fname As String
Dim ByteCt As Long
Dim bindat As Byte
Dim ct As Long
Dim str1 As String
Dim str2 As String

   On Error GoTo SkipIt

   CommonDialog1.Filter = "CPM Files|*.com||*.*|All files"
   CommonDialog1.ShowOpen
   Open CommonDialog1.FileName For Binary As #1
   Call cmdClear_Click
   ByteCt = 1
   
   While ByteCt < LOF(1)
      Get #1, ByteCt, bindat
      RAM((ByteCt + 255) And CLng("&HFFFF")) = bindat
      ByteCt = ByteCt + 1
   Wend
   Close 1
   
   txtLRegs(0).Text = "0100"
   txtStart.Text = "0100"
   Call DisplayHex
   Call DisplayAssembly
  
SkipIt:
   On Error GoTo 0
End Sub

' a HEX file has the following Format -
' 4 Hex Digits for the load address, followed by a colon and a space.
' 16 hex pairs, each followed by a space.
' example -
' 0000: 00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F
Private Sub cmdLoadHEX_Click()
Dim fname As String
Dim ByteCt As Long
Dim Ctr As Long
Dim FileBytes As Long
Dim bindat As Byte
Dim FilDat As String
Dim MemAddr As Long
Dim baseAddr As Long
Dim textLine As String
Dim DispStart As String
Dim FirstLine As Boolean
Dim rectype As Integer

   On Error GoTo SkipIt
   CommonDialog1.Filter = "HEX Files|*.hex|*.*|All files"
   CommonDialog1.ShowOpen

   Open CommonDialog1.FileName For Input As #1
   FirstLine = True
   While Not EOF(1)
      Line Input #1, textLine
      Mid(textLine, 1, 1) = " "
      textLine = Trim(textLine)
      ByteCt = val("&H" + Left(textLine, 2))
      Mid(textLine, 1, 2) = "  "
      textLine = Trim(textLine)
      baseAddr = val("&H" + Left(textLine, 4))
      Mid(textLine, 1, 4) = "    "
      textLine = Trim(textLine)
      rectype = val("&H" + Left(textLine, 2))
      Mid(textLine, 1, 2) = "  "
      textLine = Trim(textLine)
      If rectype = 0 Then
         If FirstLine Then
            DispStart = Right("0000" + Hex(baseAddr), 4)
            FirstLine = False
         End If
         For Ctr = 1 To ByteCt
            bindat = val("&H" + Left(textLine, 2))
            Mid(textLine, 1, 2) = "  "
            textLine = Trim(textLine)
            RAM((baseAddr + Ctr - 1) And CLng("&HFFFF")) = bindat
         Next Ctr
      End If
   Wend
   Close 1
   Close 2
   
   Call DisplayAssembly
   txtStart.Text = DispStart
   Call DisplayHex

SkipIt:
   On Error GoTo 0

End Sub

Private Sub cmdSaveCPM_Click()
Dim fname As String
Dim ByteCt As Long
Dim bindat As Byte
Dim ct As Long
Dim str1 As String
Dim str2 As String


   CommonDialog1.Filter = "CPM Files|*.com||*.*|All files"
   CommonDialog1.ShowSave
   Open CommonDialog1.FileName For Binary As #1
   For ByteCt = val("&H100") To val("&H" + txtSaveEnd.Text)
      bindat = RAM(ByteCt)
      Put #1, , bindat
   Next ByteCt
   Close 1
  
SkipIt:
   On Error GoTo 0

End Sub

Private Sub cmdSaveSource_Click()
Dim FileNo As Long
Dim MachCode As String
Dim SrcCode As String
Dim EndofBlock As Long

   CommonDialog1.Filter = "Assembly File|*.z80|All FIles|*.*"
   CommonDialog1.ShowSave
   If Len(CommonDialog1.FileName) > 0 Then
      FileNo = FreeFile
      Open CommonDialog1.FileName For Output As #FileNo '

      EndofBlock = CLng("&H0" + txtSaveEnd.Text)
      While Regs16.PC <= EndofBlock
      Call DissAssemble(Regs16.PC, MachCode, SrcCode)
         Print #FileNo, Left(SrcCode + String(30, " "), 30); ";"; MachCode
      Wend
      Close FileNo
   End If
End Sub

Private Sub txtOneLineASM_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SourceLine As String
Dim objstr As String
   If KeyCode = 13 Then
      
      SourceLine = txtAsmImmed.Text
      Pass2 = True
      objstr = Assemble(SourceLine)
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Left(objstr, 4)
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + " " + Left(Right(objstr, Len(objstr) - 4) + "          ", 8) + " "
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + SourceLine + vbCrLf
      txtAsmImmed.Text = ""
      ASMPC = ASMPC + PCInc
   End If
   
End Sub

Private Sub txtAsmImmed_KeyDown(KeyCode As Integer, Shift As Integer)
Dim objstr As String
Dim objaddr As Long
Dim ct As Long
    
   Pass2 = True
   Call InitVars
   If KeyCode = 13 Then
      If txtInlineAsmAddr.Text <> "" Then
         ASMPC = CLng("&H0" + txtInlineAsmAddr.Text)
         PCInc = 0
         txtInlineAsmAddr.Text = ""
      End If
      objstr = Assemble(txtAsmImmed.Text)
      frmAsmList.Show
      Load frmAsmList
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Left(objstr, 4) + "  " + Left(Right(objstr, Len(objstr) - 4) + "        ", 8) + " " + txtAsmImmed.Text + vbCrLf
      If Len(objstr) > 4 Then
         objaddr = CLng("&H0" + Left(objstr, 4))
         Mid(objstr, 1, 4) = "    "
         objstr = Trim(objstr)
         For ct = 1 To Len(objstr) Step 2
            RAM(objaddr) = val("&H0" + Mid(objstr, ct, 2))
            objaddr = (objaddr + 1) And CLng("&HFFFF")
         Next ct
      End If
      txtInlineAsmAddr.Text = Right("0000" + Hex(ASMPC + PCInc), 4)
      txtAsmImmed.Text = ""
      Call DisplayHex
      Call DisplayAssembly
      
   End If
   Pass2 = False
   txtAsmImmed.SetFocus
End Sub

Private Sub cmdAssembleFile_Click()
Dim SourceLine As String
Dim tempstr As String
Dim ct As Long
Dim objstr As String
Dim objstr1 As String
Dim objaddr As Long

   frmAsmList.Show
   Load frmAsmList
   CPMSel = optCPMcom
   BINSel = optBINbin
   HexSel = optIntelHex
   
   Call InitVars
   EndOp = False
   frmAsmList.txtAsmList.Text = ""
   If optCPMcom Then
      ASMPC = &H100
   Else
      ASMPC = 0
   End If
   Pass2 = False
   IntelHexCt = 0
   SrcNum = FreeFile
   Open txtAsmSource.Text For Input As #SrcNum
   While Not EOF(1)
      Line Input #SrcNum, SourceLine
      Assemble (SourceLine)
   Wend
   Close #SrcNum
   
   If optCPMcom Then
      ASMPC = &H100
   Else
      ASMPC = 0
   End If
   Pass2 = True
   
   SrcNum = FreeFile
   Open txtAsmSource.Text For Input As #SrcNum
   ListNum = FreeFile
   Open txtAsmList.Text For Output As #ListNum
   ObjectNum = FreeFile
   If optIntelHex Then
      Open txtAsmObj.Text For Output As #ObjectNum
   ElseIf optCPMcom Then
      Open txtAsmObj.Text For Binary As #ObjectNum
   End If
   
   Print #ListNum, txtAsmSource.Text
   Call InitObject
   While Not EOF(SrcNum) And (Not EndOp)
      Line Input #SrcNum, SourceLine
      Pass2 = True
      objstr = Assemble(SourceLine)
      objstr1 = objstr
      If (Len(objstr) > 4) And (InStr(1, objstr, "=") = 0) Then Call ObjectSave(objstr)
      Print #2, Left(objstr, 4) & " ";
      Print #2, Left(Right(objstr, Len(objstr) - 4) + "          ", 8);
      Print #2, "  " & SourceLine + vbCrLf;
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Left(objstr, 4) + "  "
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Left(Right(objstr, Len(objstr) - 4) + "        ", 8)
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + " " + SourceLine + vbCrLf
      Mid(objstr, 1, 4) = String(4, " ")
      objstr = Trim(objstr)
      If Len(objstr) > 8 Then
         While Len(objstr) > 0
            Mid(objstr, 1, 8) = "        "
            objstr = Trim(objstr)
            If Len(objstr) > 0 Then
               Print #2, "      " + Left(objstr, 8)
               frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + "      "
               frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Left(objstr + "          ", 8) + vbCrLf
            End If
         Wend
      End If
      If ErrorStr <> "" Then
         Print #ListNum, ErrorStr
         frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + ErrorStr + vbCrLf
      End If
      'DoEvents
      ErrorStr = ""
           
      If Len(objstr1) > 4 Then
         objaddr = CLng("&H0" + Left(objstr1, 4))
         Mid(objstr1, 1, 4) = "    "
         objstr1 = Trim(objstr1)
         If Left(objstr1, 1) <> "*" Then
            For ct = 1 To Len(objstr1) Step 2
               RAM(objaddr) = val("&H0" + Mid(objstr1, ct, 2))
               objaddr = (objaddr + 1) And CLng("&HFFFF")
            Next ct
         Else
          txtLRegs(0).Text = Right(objstr1, 4)
          txtStart.Text = Right(objstr1, 4)
         End If
      End If

   Wend
   ObjectClose (objstr1)
   frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + vbCrLf + vbCrLf + "Symbol Table" + vbCrLf
   Print #ListNum,
   Print #ListNum, "Symbol Table"
   For ct = 1 To SymTabLast - 1
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Left(symTable(ct).name + "                  ", 16)
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + Right("0000" + Hex(symTable(ct).value), 4) + "  "
      If Not symTable(ct).defined Then frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + "Undefined  "
      If symTable(ct).multiDef Then frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + "Multple Defined  "
      frmAsmList.txtAsmList.Text = frmAsmList.txtAsmList.Text + vbCrLf
      Print #ListNum, Left(symTable(ct).name + "                  ", 16);
      Print #ListNum, Right("0000" + Hex(symTable(ct).value), 4) + "  ";
      If Not symTable(ct).defined Then Print #2, "Undefined  ";
      If symTable(ct).multiDef Then Print #2, "Multple Defined  ";
      Print #ListNum, " "
      Memory(symTable(ct).value).Label = symTable(ct).name
      
   Next ct

   Close SrcNum
   Close ListNum
   Close ObjectNum
   
   Call DisplayHex
   Call DisplayAssembly
   
End Sub

Sub DisplayRegs()

   Regs8.F = 0
   lblSflg.Caption = "0"
   lblZflg.Caption = "0"
   lblHflg.Caption = "0"
   lblPflg.Caption = "0"
   lblNflg.Caption = "0"
   lblCflg.Caption = "0"
   If Flags.S Then
      Regs8.F = Regs8.F + 128
      lblSflg.Caption = "1"
   End If
   If Flags.Z Then
      Regs8.F = Regs8.F + 64
      lblZflg.Caption = "1"
   End If
   If Flags.H Then
      Regs8.F = Regs8.F + 16
      lblHflg.Caption = "1"
   End If
   If Flags.P Then
      Regs8.F = Regs8.F + 4
      lblPflg.Caption = "1"
   End If
   If Flags.N Then
      Regs8.F = Regs8.F + 2
      lblNflg.Caption = "1"
   End If
   If Flags.C Then
      Regs8.F = Regs8.F + 1
      lblCflg.Caption = "1"
   End If

   Regs8.S = (Regs16.SP And CLng(65280)) / 256
   Regs8.P = Regs16.SP And &HFF
   
   txtRegs(0).Text = Right("00" + Hex(Regs8.A), 2)
   txtRegs(1).Text = Right("00" + Hex(Regs8.F), 2)
   txtRegs(2).Text = Right("00" + Hex(Regs8.B), 2)
   txtRegs(3).Text = Right("00" + Hex(Regs8.C), 2)
   txtRegs(4).Text = Right("00" + Hex(Regs8.D), 2)
   txtRegs(5).Text = Right("00" + Hex(Regs8.E), 2)
   txtRegs(6).Text = Right("00" + Hex(Regs8.H), 2)
   txtRegs(7).Text = Right("00" + Hex(Regs8.L), 2)
   txtRegs(8).Text = Right("00" + Hex(AltRegs8.A), 2)
   txtRegs(9).Text = Right("00" + Hex(AltRegs8.F), 2)
   txtRegs(10).Text = Right("00" + Hex(AltRegs8.B), 2)
   txtRegs(11).Text = Right("00" + Hex(AltRegs8.C), 2)
   txtRegs(12).Text = Right("00" + Hex(AltRegs8.D), 2)
   txtRegs(13).Text = Right("00" + Hex(AltRegs8.E), 2)
   txtRegs(14).Text = Right("00" + Hex(AltRegs8.H), 2)
   txtRegs(15).Text = Right("00" + Hex(AltRegs8.L), 2)
   txtRegs(16).Text = Right("00" + Hex(Regs8.I), 2)
   txtRegs(17).Text = Right("00" + Hex(Regs8.R), 2)
   txtLRegs(0).Text = Right("0000" + Hex(Regs16.PC), 4)
   txtLRegs(1).Text = Right("0000" + Hex(Regs16.SP), 4)
   txtLRegs(2).Text = Right("0000" + Hex(Regs16.IX), 4)
   txtLRegs(3).Text = Right("0000" + Hex(Regs16.IY), 4)
End Sub

Sub DisplayHex()
Dim addr As Long
Dim ct As Long
Dim Offset As Long
Dim workStr As String
Dim AsciiStr As String

   txtHexdisp.Text = ""
   workStr = Right("0000" + txtStart.Text, 4)
   addr = CLng("&H0" + workStr)
   addr = addr And CLng("&HFFFF")
   Offset = 16
   txtHexdisp.Text = "      0  1  2  3  4  5  6  7   8  9  A  B  C  D  E  F |01234567 89ABCDEF|" + vbCrLf
   For ct = addr To addr + 255
      If Offset >= 16 Then
         txtHexdisp.Text = txtHexdisp.Text + Right("0000" + Hex(ct), 4) + ":"
         Offset = 0
         AsciiStr = ""
      End If
      txtHexdisp.Text = txtHexdisp.Text + Right("00" + Hex(RAM(ct And CLng("&HFFFF"))), 2) + " "
      If RAM(ct And CLng("&HFFFF")) > 32 Then
         AsciiStr = AsciiStr + Chr(RAM(ct And CLng("&HFFFF")))
      Else
         AsciiStr = AsciiStr + " "
      End If
      If Offset = 7 Then
         txtHexdisp.Text = txtHexdisp.Text + " "
         AsciiStr = AsciiStr + " "
      End If
      Offset = Offset + 1
      If Offset >= 16 Then txtHexdisp.Text = txtHexdisp.Text + "|" + AsciiStr + "|" + vbCrLf
   Next ct
          
End Sub

Sub DisplayAssembly()
Dim NumLines As Integer
Dim lines As Integer
Dim MachCode As String
Dim SrcCode As String
Dim PC As Long

   txtDissasm.Text = ""
   NumLines = txtDissasm.Height \ 245
   lines = 1
   OperMode = 3
   PC = Regs16.PC
   While lines < NumLines
      Call DissAssemble(PC, MachCode, SrcCode)
      lines = lines + 1
      txtDissasm.Text = txtDissasm.Text + Left(MachCode + "              ", 14) + SrcCode + vbCrLf
   Wend

End Sub


Private Sub cmdClear_Click()
Dim ct As Long
   For ct = 0 To CLng("&HFFFF")
      RAM(ct) = 0
      Memory(ct).Label = ""
      Memory(ct).Usage = 0
   Next ct
   
   Call DisplayHex
   Call DisplayAssembly

End Sub


Private Sub cmdSelSOurce_Click()
   
   CommonDialog1.Filter = "z80 Assembly File|*.z80|Assembly File|*.asm|All FIles|*.*"
   CommonDialog1.ShowOpen
   
   If Len(CommonDialog1.FileName) > 0 Then
      txtAsmSource.Text = CommonDialog1.FileName
      txtAsmList.Text = Left(txtAsmSource.Text, Len(txtAsmSource.Text) - 4) + ".LST"
      If optIntelHex.value Then
         txtAsmObj.Text = Left(txtAsmSource.Text, Len(txtAsmSource.Text) - 4) + ".HEX"
      End If
      If optCPMcom.value Then
         txtAsmObj.Text = Left(txtAsmSource.Text, Len(txtAsmSource.Text) - 4) + ".com"
      End If
      cmdAssembleFile.Visible = True
   End If
   frmAsmList.txtAsmList.Text = ""

End Sub


Private Sub Form_Resize()

   If frmz80Emu.Height < 500 Then Exit Sub
   If frmz80Emu.Width < 500 Then Exit Sub
   If frmz80Emu.Height < 9700 Then frmz80Emu.Height = 9700
   txtDissasm.Height = frmz80Emu.Height - 3300
   If frmz80Emu.Width < 15500 Then frmz80Emu.Width = 15500
   Call DisplayAssembly

End Sub

Private Sub txtHexDisp_KeyPress(KeyAscii As Integer)
Dim str1, str2 As String
Dim PosHold As Long
Dim CursCol, CursRow As Long
Dim collAddr As Long
Dim nibl As Long
Dim Byteval As Byte
Dim ByteAddr As Long
Dim NiblVal As Byte
Dim mask As Byte
   
   Updatelabels = False
   
   PosHold = txtHexdisp.SelStart
   If (txtHexdisp.SelStart \ 75) < 1 Then
      KeyAscii = 0
      Call DisplayHex
      
   Else
      CursRow = txtHexdisp.SelStart \ 75
      CursCol = txtHexdisp.SelStart Mod 75

      If (CursCol > 4) And (CursCol < 53) Then
         If CursCol > 28 Then
            collAddr = CursCol - 1
         Else
            collAddr = CursCol
         End If
         nibl = (collAddr - 5) Mod 3
         If nibl = 2 Then KeyAscii = 32
         collAddr = (collAddr - 5) \ 3
         NiblVal = val("&H" + Chr(KeyAscii))
         ByteAddr = (val("&H1" + Mid(txtHexdisp.Text, CursRow * 75 + 1, 4)) + collAddr) And CLng("&HFFFF")
         If (KeyAscii = 48) Or ((NiblVal > 0) And (NiblVal < 16)) Or ((nibl = 2) And (KeyAscii = 32)) Then
            str1 = Left(txtHexdisp.Text, txtHexdisp.SelStart)
            str2 = Right(txtHexdisp.Text, Len(txtHexdisp.Text) - (txtHexdisp.SelStart + 1))
            txtHexdisp.Text = str1 + UCase(Chr(KeyAscii)) + str2
            KeyAscii = 0
            If nibl = 0 Then
               txtHexdisp.SelStart = PosHold + 1
            ElseIf CursCol = 27 Then
               txtHexdisp.SelStart = PosHold + 3
            ElseIf nibl = 1 Then
               txtHexdisp.SelStart = PosHold + 2
            ElseIf CursCol = 30 Then
               txtHexdisp.SelStart = PosHold + 3
            End If
            If nibl = 0 Then
               mask = &HF
               NiblVal = NiblVal * 16
            Else
               mask = &HF0
            End If
            RAM(ByteAddr) = (RAM(ByteAddr) And mask) Or NiblVal
            Call DisplayAssembly
         Else
            KeyAscii = 0
         End If
      Else
         KeyAscii = 0
      End If
   End If
   If (CursCol > 51) And (CursRow < 16) Then
      txtHexdisp.SelStart = (CursRow + 1) * 75 + 5
   ElseIf (CursCol > 51) And (CursRow = 16) Then
      txtStart.Text = Right("0000" + (Hex(val("&H" + txtStart.Text) + 16)), 4)
      txtHexdisp.SelStart = 16 * 75 + 5
   End If
   Call DisplayRegs
End Sub

Private Sub txtInlineAsmAddr_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtAsmImmed.SetFocus
   End If
End Sub



Private Sub txtLRegs_Change(Index As Integer)
Dim RegHold As Long
   If Len(txtLRegs(Index).Text) < 4 Then Exit Sub
   RegHold = CLng("&H" + txtLRegs(Index).Text)
   Select Case Index
      Case 0: Regs16.PC = RegHold
              Call DisplayAssembly
      Case 1: Regs16.SP = RegHold
      Case 2: Regs16.IX = RegHold
      Case 3: Regs16.IY = RegHold
   End Select

End Sub

Private Sub txtLRegs_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RegPC As Long
   
   If Index > 0 Then Exit Sub
   
   If ((KeyCode = 33) Or (KeyCode = 34) Or (KeyCode = 38) Or (KeyCode = 40)) Then
      RegPC = CLng("&H" + txtLRegs(0).Text)
      If Shift Then KeyCode = KeyCode + &H100
      Select Case KeyCode
         Case &H21: RegPC = RegPC + &H10    ' <page up>
         Case &H22: RegPC = RegPC - &H10    ' <page down>
         Case &H26: RegPC = RegPC + 1        ' <up arrow>
         Case &H28: RegPC = RegPC - 1        ' <down arrow>
         Case &H121: RegPC = RegPC + &H100  ' <shift><page Up>
         Case &H122: RegPC = RegPC - &H100  ' <shift><page down>
      End Select
      txtLRegs(0).Text = Right("000" + Hex(RegPC), 4)
      
   End If

End Sub

Private Sub txtRegs_Change(Index As Integer)
Dim RegHold As Long
   If Len(txtRegs(Index).Text) < 2 Then Exit Sub
   RegHold = CLng("&H" + txtRegs(Index).Text)
   Select Case Index
      Case 0: Regs8.A = RegHold
      Case 1: Regs8.F = RegHold
      Case 2: Regs8.B = RegHold
      Case 3: Regs8.C = RegHold
      Case 4: Regs8.D = RegHold
      Case 5: Regs8.E = RegHold
      Case 6: Regs8.H = RegHold
      Case 7: Regs8.L = RegHold
      Case 8: Regs8.A = RegHold
      Case 9: Regs8.F = RegHold
      Case 10: AltRegs8.B = RegHold
      Case 11: AltRegs8.C = RegHold
      Case 12: AltRegs8.D = RegHold
      Case 13: AltRegs8.E = RegHold
      Case 14: AltRegs8.H = RegHold
      Case 15: AltRegs8.L = RegHold
      Case 16: AltRegs8.I = RegHold
      Case 17: AltRegs8.R = RegHold
   End Select
  
End Sub

Private Sub txtSaveEnd_Change()
   txtSaveEnd.Text = UCase(Trim(txtSaveEnd.Text))
   If txtSaveEnd.Text <> "" Then
      cmdSaveBin.Enabled = True
      cmdSaveCPM.Enabled = True
      cmdSaveSource.Enabled = True
   End If
   If (txtSaveStart.Text <> "") And (txtSaveEnd.Text <> "") Then
      cmdSaveHex.Enabled = True
   End If
   
End Sub

Private Sub txtSaveStart_Change()
   txtSaveStart.Text = UCase(Trim(txtSaveStart.Text))
   If (txtSaveStart.Text <> "") And (txtSaveEnd.Text <> "") Then
      cmdSaveCPM.Enabled = True
      cmdSaveBin.Enabled = True
      cmdSaveHex.Enabled = True
   End If
   
End Sub


Private Sub txtStart_Change()
   txtStart.Text = UCase(txtStart.Text)
   Call DisplayHex

End Sub


Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
Dim HexStart As Long

      If ((KeyCode = 33) Or (KeyCode = 34) Or (KeyCode = 38) Or (KeyCode = 40)) Then
      HexStart = CLng("&H" + txtStart.Text)
      If Shift Then KeyCode = KeyCode + &H100
      Select Case KeyCode
         Case &H21: HexStart = HexStart + &H100
         Case &H22: HexStart = HexStart - &H100
         Case &H26: HexStart = HexStart + &H10
         Case &H28: HexStart = HexStart - &H10
         Case &H121: HexStart = HexStart + &H1000
         Case &H122: HexStart = HexStart - &H1000
      End Select
      If Len(txtStart.Text) > 3 Then txtStart.Text = Right("0000" + Hex(HexStart), 4)
      
   End If
End Sub

Private Sub txtCtrlAdd_Change()
   txtCtrlWord.Text = ""
   If Len(txtCtrlAdd.Text) = 4 Then txtCtrlWord.SetFocus
   
End Sub

Private Sub txtCtrlWord_Change()
Dim ctrlAdd As Long
Dim ct As Long
   
   txtCtrlWord.Text = UCase(txtCtrlWord.Text)
   ctrlAdd = CLng("&H" + txtCtrlAdd.Text)
   Select Case txtCtrlWord.Text
      Case "A": Memory(ctrlAdd).Usage = 1
      Case "B": Memory(ctrlAdd).Usage = 2
      Case "I": Memory(ctrlAdd).Usage = 3
      Case "S": Memory(ctrlAdd).Usage = 4
      Case "W": Memory(ctrlAdd).Usage = 5
      Case "X":
                If Memory(ctrlAdd).Usage = 6 Then
                ct = 0
                   For ct = 1 To CLng("&HFFFF")
                      If Memory(ct).Usage = 6 Then
                         Memory(ct).Usage = 0
                      End If
                   Next ct
                   EndAddress = CLng("&H10000")
                Else
                   Memory(ctrlAdd).Usage = 0
                End If
      Case "E": Memory(ctrlAdd).Usage = 6
                Memory(ctrlAdd + 1).Usage = 6
                Memory(ctrlAdd + 2).Usage = 6
                Memory(ctrlAdd + 3).Usage = 6
                Memory(ctrlAdd + 4).Usage = 6
      Case Else:
   End Select
   Call DisplayAssembly
   txtCtrlAdd.SetFocus
End Sub

Private Sub txtLabelAdd_Change()
   txtLabelName.Text = ""
   If Len(txtLabelAdd.Text) = 4 Then txtLabelName.SetFocus
End Sub

Private Sub txtLabelName_KeyDown(KeyCode As Integer, Shift As Integer)
Dim address As Long

   If KeyCode = 13 Then
      address = CLng("&H0" + txtLabelAdd.Text)
      Memory(address).Label = txtLabelName.Text
      KeyCode = 0
      txtLabelName.Text = ""
      Call DisplayAssembly
      txtLabelAdd.SetFocus
   End If
End Sub
