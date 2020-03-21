VERSION 5.00
Begin VB.Form frmConsoleDisplay 
   Caption         =   "Z80 Emulator Console"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConsole 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6645
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmConsoleDisplay.frx":0000
      Top             =   30
      Width           =   12105
   End
End
Attribute VB_Name = "frmConsoleDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

