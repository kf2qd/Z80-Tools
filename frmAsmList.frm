VERSION 5.00
Begin VB.Form frmAsmList 
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAsmList 
      BeginProperty Font 
         Name            =   "Dina"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7300
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmAsmList.frx":0000
      Top             =   75
      Width           =   7200
   End
End
Attribute VB_Name = "frmAsmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   txtAsmList.Text = ""
End Sub

Private Sub Form_Resize()
   If frmAsmList.Height < 4000 Then frmAsmList.Height = 4000
   If frmAsmList.Width < 5000 Then frmAsmList.Width = 5000
   txtAsmList.Height = frmAsmList.Height - 700
   txtAsmList.Width = frmAsmList.Width - 300
   
End Sub
