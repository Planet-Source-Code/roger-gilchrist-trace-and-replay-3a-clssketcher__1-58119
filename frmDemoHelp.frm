VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDemoHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demo Help"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8070
      _Version        =   393217
      ScrollBars      =   2
      FileName        =   "C:\DownLoads2005\UNDERDEVELOPMENT\ClsSketcher\Release_5-01-2005_12-31-42 PM\Release_6-01-2005_1-13-26 AM\ClsSketcher Demo Help.rtf"
      TextRTF         =   $"frmDemoHelp.frx":0000
   End
   Begin VB.CommandButton cmdHelpClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   255
      Left            =   7560
      TabIndex        =   0
      Top             =   4680
      Width           =   975
   End
End
Attribute VB_Name = "frmDemoHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'just a very stupid help form
Private Sub cmdHelpClose_Click()

  Unload Me

End Sub

Private Sub Form_Load()

  Me.Caption = "Trace And Replay 2 (ClsSketcher) Demo Help"
 
End Sub


':)Code Fixer V2.8.3 (6/01/2005 7:58:16 AM) 1 + 21 = 22 Lines Thanks Ulli for inspiration and lots of code.

