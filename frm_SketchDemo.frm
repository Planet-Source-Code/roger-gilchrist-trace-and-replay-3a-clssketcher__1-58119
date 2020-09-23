VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_SketchDemo 
   Caption         =   "Trace And Replay"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9900
   Icon            =   "frm_SketchDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdlDemo 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblHint 
      Caption         =   "To draw, click and drag mouse pointer. (Right-button auto-draws back to start point)"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpt 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "&Open (Instant)"
         Index           =   1
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "Open (&Timed)"
         Index           =   2
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "&Reload"
         Index           =   3
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "Load UnderLay Image"
         Index           =   5
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "&Save( Sketch)"
         Index           =   7
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "Save &As ( Sketch)"
         Index           =   8
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "Save( &BMP)"
         Index           =   9
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFileOpt 
         Caption         =   "E&xit"
         Index           =   11
      End
   End
   Begin VB.Menu mnuUnderlay 
      Caption         =   "UnderLay"
      Begin VB.Menu mnuUnderlayOpt 
         Caption         =   "Hide/Show"
         Index           =   0
      End
      Begin VB.Menu mnuUnderlayOpt 
         Caption         =   "Offset "
         Index           =   1
         Begin VB.Menu mnuOffsetOpt 
            Caption         =   "None"
            Index           =   0
         End
         Begin VB.Menu mnuOffsetOpt 
            Caption         =   "Right"
            Index           =   1
         End
         Begin VB.Menu mnuOffsetOpt 
            Caption         =   "Down"
            Index           =   2
         End
         Begin VB.Menu mnuOffsetOpt 
            Caption         =   "Left"
            Index           =   3
         End
         Begin VB.Menu mnuOffsetOpt 
            Caption         =   "Up"
            Index           =   4
         End
         Begin VB.Menu mnuOffsetOpt 
            Caption         =   "Right/Down"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuColour 
      Caption         =   "&Colour"
   End
   Begin VB.Menu mnuWidth 
      Caption         =   "&Width(1)"
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "1"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "2"
         Index           =   1
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "3"
         Index           =   2
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "4"
         Index           =   3
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "5"
         Index           =   4
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "6"
         Index           =   5
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "7"
         Index           =   6
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "8"
         Index           =   7
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "9"
         Index           =   8
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "10"
         Index           =   9
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "11"
         Index           =   10
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "12"
         Index           =   11
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "13"
         Index           =   12
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "14"
         Index           =   13
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "15"
         Index           =   14
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "16"
         Index           =   15
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "17"
         Index           =   16
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "18"
         Index           =   17
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "19"
         Index           =   18
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "20"
         Index           =   19
      End
      Begin VB.Menu mnuWidthOpt 
         Caption         =   "Other (Use F2 to set)"
         Index           =   20
      End
   End
   Begin VB.Menu mnuPen 
      Caption         =   "&Pen (Standard)"
      Begin VB.Menu mnuPenOpt 
         Caption         =   "Standard"
         Index           =   0
      End
      Begin VB.Menu mnuPenOpt 
         Caption         =   "Duplicate"
         Index           =   1
      End
   End
   Begin VB.Menu mnuUndo 
      Caption         =   "&Undo"
      Begin VB.Menu mnuUndoOpt 
         Caption         =   "Undo 1 (BackSpace)"
         Index           =   0
      End
      Begin VB.Menu mnuUndoOpt 
         Caption         =   "ReLoad"
         Index           =   1
      End
      Begin VB.Menu mnuUndoOpt 
         Caption         =   "UnDo All"
         Index           =   2
      End
   End
   Begin VB.Menu mnuRedraw 
      Caption         =   "&ReDraw"
      Begin VB.Menu mnuRedDrawOpt 
         Caption         =   "&Clear (Sket Hidden)"
         Index           =   0
      End
      Begin VB.Menu mnuRedDrawOpt 
         Caption         =   "Instant"
         Index           =   1
      End
      Begin VB.Menu mnuRedDrawOpt 
         Caption         =   "Timed"
         Index           =   2
         Begin VB.Menu mnuSpeed 
            Caption         =   "User Speed (F3/Shift+F3)"
            Index           =   0
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Very Slow (60)"
            Index           =   1
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Slow         (30)"
            Index           =   2
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Medium    (15)"
            Index           =   3
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Fast          (5)"
            Index           =   4
         End
         Begin VB.Menu mnuSpeed 
            Caption         =   "Very Fast  (1)"
            Index           =   5
         End
      End
      Begin VB.Menu mnuRedDrawOpt 
         Caption         =   "Scale"
         Index           =   3
         Begin VB.Menu mnuScale 
            Caption         =   "10%"
            Index           =   0
         End
         Begin VB.Menu mnuScale 
            Caption         =   "25%"
            Index           =   1
         End
         Begin VB.Menu mnuScale 
            Caption         =   "50%"
            Index           =   2
         End
         Begin VB.Menu mnuScale 
            Caption         =   "75%"
            Index           =   3
         End
         Begin VB.Menu mnuScale 
            Caption         =   "100%"
            Index           =   4
         End
         Begin VB.Menu mnuScale 
            Caption         =   "150%"
            Index           =   5
         End
         Begin VB.Menu mnuScale 
            Caption         =   "200%"
            Index           =   6
         End
         Begin VB.Menu mnuScale 
            Caption         =   "300%"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuAuto 
      Caption         =   "Auto"
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Horz-Line"
         Index           =   0
      End
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Vert-Line"
         Index           =   1
      End
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Horz-Rough"
         Index           =   2
      End
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Vert-Rough"
         Index           =   3
      End
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Point Dot"
         Index           =   4
      End
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Point Rain"
         Index           =   5
      End
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Point Dash"
         Index           =   6
      End
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Point Dash Interactive"
         Index           =   7
      End
      Begin VB.Menu mnuAutoOpt 
         Caption         =   "Random"
         Index           =   8
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frm_SketchDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Skt    As New ClsSketcher
Private HideLabel As Long
Private Sub Form_Load()

  Skt.Initialize Me, cdlDemo
  mnuUnderlay.Enabled = Skt.UsingUnderLay
  mnuAuto.Enabled = Skt.UsingUnderLay

End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If HideLabel < 40 Then
  HideLabel = HideLabel + 1
  Else
  If lblHint.Visible Then ' turn of the hint label
    lblHint.Visible = False
  End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)


End Sub

Private Sub Form_Unload(Cancel As Integer)

  End

End Sub

Private Sub mnuAutoOpt_Click(Index As Integer)

  Select Case Index
   Case 0
    Skt.Auto_Sketch eHorzLine
   Case 1
    Skt.Auto_Sketch eVertLine
   Case 2
    Skt.Auto_Sketch eRoughH
   Case 3
    Skt.Auto_Sketch eRoughV
   Case 4
    Skt.Auto_Sketch ePointDot
   Case 5
    Skt.Auto_Sketch ePointRain
   Case 6
    Skt.Auto_Sketch ePointDash
   Case 7
    Skt.Auto_Sketch ePointInteract
   Case 8
    Skt.Auto_Sketch eRandom
  End Select

End Sub

Private Sub mnuColour_Click()

  Skt.DoColorDlg

End Sub

Private Sub mnuFile_Click()

  mnuFileOpt(3).Enabled = Skt.HasData
  mnuFileOpt(7).Enabled = Skt.HasData
  mnuFileOpt(8).Enabled = Skt.HasData
  mnuFileOpt(9).Enabled = Skt.HasData

End Sub

Private Sub mnuFileOpt_Click(Index As Integer)

  Select Case Index
   Case 0 'New
    Skt.NewPic
   Case 1, 2 'Open
    If lblHint.Visible Then ' turn of the hint label
      lblHint.Visible = False
    End If
    Skt.LoadSket Index = 2
    mnuWidthOpt_Click Skt.CurrentWidth
   Case 3
    Skt.ReLoad
'Case 4'--
   Case 5
    Skt.LoadUnderLay
    mnuUnderlay.Enabled = Skt.UsingUnderLay
    mnuAuto.Enabled = Skt.UsingUnderLay
    mnuUnderlayOpt(0).Caption = "Hide"
    Skt.RedrawEngine
'   Case 6 '-
   Case 7, 8
    Skt.SaveSket Index = 8
   Case 9
    Skt.SaveBMP
'Case 10 '---
   Case 11 'Exit
    Unload Me
  End Select

End Sub

Private Sub mnuHelp_Click()

  frmDemoHelp.Show

End Sub

Private Sub mnuOffsetOpt_Click(Index As Integer)

  MenuChecking mnuOffsetOpt, Index
  Select Case Index
   Case 0
    Skt.TraceOffSetX = eOverLay
    Skt.TraceOffSetY = eOverLay
   Case 1
    Skt.TraceOffSetX = ePositive
   Case 2
    Skt.TraceOffSetY = ePositive
   Case 3
    Skt.TraceOffSetX = eNegative
   Case 4
    Skt.TraceOffSetY = eNegative
   Case 5
    Skt.TraceOffSetX = ePositive
    Skt.TraceOffSetY = ePositive
  End Select
  Skt.RedrawEngine

End Sub

Private Sub mnuPenOpt_Click(Index As Integer)

  MenuChecking mnuPenOpt, Index
  Skt.PenMode = Index
  mnuPen.Caption = "&Pen(" & mnuPenOpt(Index).Caption & ")"

End Sub

Private Sub mnuRedDrawOpt_Click(Index As Integer)

  Select Case Index
   Case 0 'clear drawing
    Skt.Cls
   Case 1 'Instant
    Skt.RedrawEngine
'case 2'Timed sub-menu
  End Select

End Sub

Private Sub mnuScale_Click(Index As Integer)

  MenuChecking mnuScale, Index
  Select Case Index
   Case 0 '10
    Skt.CurrentScale = 0.1
   Case 1 '25
    Skt.CurrentScale = 0.25
   Case 2 '50
    Skt.CurrentScale = 0.5
   Case 3 '75
    Skt.CurrentScale = 0.75
   Case 4 '100
    Skt.CurrentScale = 1
   Case 5 '150
    Skt.CurrentScale = 1.5
   Case 6 '200
    Skt.CurrentScale = 2
   Case 7 '300
    Skt.CurrentScale = 3
  End Select
  Skt.RedrawEngine

End Sub

Private Sub mnuSpeed_Click(Index As Integer)

'timed redrawing

  Select Case Index
'Case 0 ' just use the class's current speed setting
   Case 1
    Skt.Speed = 60
   Case 2
    Skt.Speed = 30
   Case 3
    Skt.Speed = 15
   Case 4
    Skt.Speed = 5
   Case 5
    Skt.Speed = 1
  End Select
  Skt.RedrawEngine True

End Sub

Private Sub mnuUnderlay_Click()

  mnuUnderlayOpt(0).Enabled = Skt.UsingUnderLay

End Sub

Private Sub mnuUnderlayOpt_Click(Index As Integer)

  If Index = 0 Then
    mnuUnderlayOpt(0).Caption = IIf(Skt.UnderLayVisible, "Show", "Hide")
    Skt.UnderLayToggle
    Skt.RedrawEngine
  End If

End Sub

Private Sub mnuUndoOpt_Click(Index As Integer)

  Select Case Index
   Case 0
    Skt.Undo1
   Case 1
    Skt.ReLoad
   Case 2
    Skt.NewPic
  End Select

End Sub

Private Sub mnuWidthOpt_Click(Index As Integer)

  If Index < 20 Then
    MenuChecking mnuWidthOpt, Index
    Skt.CurrentWidth = Index + 1
    mnuWidth.Caption = "&Width(" & Index + 1 & ")"
   Else
    MenuChecking mnuWidthOpt, 20
    mnuWidth.Caption = "&Width(" & Skt.CurrentWidth & ")"
  End If

End Sub

':)Code Fixer V2.8.3 (6/01/2005 7:58:16 AM) 2 + 236 = 238 Lines Thanks Ulli for inspiration and lots of code.

