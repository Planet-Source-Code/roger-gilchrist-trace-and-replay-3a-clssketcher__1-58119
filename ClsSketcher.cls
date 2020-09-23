VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSketcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Apologies for lack of comments
'This release is to get feedback/suggestions
'Please report bugs or make suggestions
'
'Idea suggested By 'Retrace your Footsteps!' at
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=58084&lngWId=1
Public Enum eAutoMode
  eHorzLine
  eVertLine
  eRoughH
  eRoughV
  ePointDot
  ePointRain
  ePointDash
  ePointInteract
  eRandom
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eHorzLine, eVertLine, eRoughH, eRoughV, ePointDot, ePointRain, ePointDash, ePointInteract, eRandom
#End If
Public Enum eOffSet
  eNegative
  eOverLay
  ePositive
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eNegative, eOverLay, ePositive
#End If
Public Enum ePen
  eStandard
  eDuplicate
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eStandard, eDuplicate
#End If
Private m_PenMode                   As ePen
Public Enum eDmode
  eNone
  eStart
  eDraw
  EEnd
  EAuto
End Enum
#If False Then
Private eNone, eStart, eDraw, EEnd, EAuto
#End If
Private Const strDefaultFile        As String = "SketchDefault.Sket"
Private m_Colour                    As Long
Private m_Wid                       As Long
Private Type DrawData
  SX                                As Long
  SY                                As Long
  SMode                             As eDmode
  sCol                              As Long
  sWid                              As Long
End Type
Private StartX                      As Long
Private StartY                      As Long
Private Sketch()                    As DrawData
Private m_Cdlg                      As CommonDialog
Private bFileAccess                 As Boolean    ' disables cdlg stuff if cdlg is not initialized
Public WithEvents m_Canvas          As Form
Attribute m_Canvas.VB_VarHelpID = -1
Private m_BTimed                    As Boolean
Private m_Scale                     As Single
Private strBaseCaption              As String
Private bAbort                      As Boolean
Private bSaving                     As Boolean
Private OffsetX                     As Long
Private OffsetY                     As Long
Private bDirty                      As Boolean
Private m_Speed                     As Long
Private bWorking                    As Boolean
Private strSketchFile               As String
Private SketchFileNameOnly          As String
Private strUnderlayFile             As String
Private bUnderLayToggleVisible      As Boolean
Private m_UsingUnderLay             As Boolean
Private DrawingPoint                As Long
Private arrColourTable              As Variant
Private arrQuickColourTable              As Variant

Public Sub Auto_Sketch(mode As eAutoMode)

  Dim InitWid As Long
  Dim Rj      As Long
  Dim RI      As Long
  Dim Rj2     As Long
  Dim RI2     As Long
  Dim I       As Long
  Dim J       As Long
  Dim J2      As Long
  Dim I2      As Long

  PenMode = eDuplicate
  J2 = m_Canvas.ScaleX(m_Canvas.Picture.Width)
  I2 = m_Canvas.ScaleY(m_Canvas.Picture.Height)
  If m_Canvas.Picture.Width > m_Canvas.Width / 2 Then
    TraceOffSetX = ePositive
    TraceOffSetY = eOverLay
   Else
    TraceOffSetX = eOverLay
    TraceOffSetY = ePositive
  End If
  If mode = eRandom Then
    mode = Int(Rnd(eRandom))
  End If
  m_Canvas.MousePointer = vbHourglass
  m_BTimed = True
  Select Case mode
   Case eHorzLine
    DrawPoint 1, 1, eStart
    CurrentWidth = 10
    For I = 1 To I2 Step 10
      DrawPoint 0, I, eStart
      For J = 1 To J2 Step 10
        CurrentWidth = Rnd * 5 + 10
        DrawPoint J, I, eDraw
        DoEvents
        If Not m_BTimed Then
          Exit For
        End If
      Next J
      DrawPoint J2, I, EEnd
      If Not m_BTimed Then
        Exit For
      End If
    Next I
   Case eVertLine
    DrawPoint 1, 1, eStart
    CurrentWidth = 10
    For J = 1 To J2 Step 10
      DrawPoint J, 0, eStart
      For I = 1 To I2 Step 10
        CurrentWidth = Rnd * 5 + 10
        DrawPoint J, I, eDraw
        DoEvents
        If Not m_BTimed Then
          Exit For
        End If
      Next I
      DrawPoint J, I2, EEnd
      If Not m_BTimed Then
        Exit For
      End If
    Next J
   Case eRoughH
    DrawPoint 1, 1, eStart
    CurrentWidth = 10
    For I = 1 To I2 Step 10
      DrawPoint 0, I, eStart
      For J = 1 To J2 Step 10
        CurrentWidth = Rnd * 5 + 10
        DrawPoint J + Rnd * 5 - 10, I + Rnd * 5 - 10, eDraw
        DoEvents
        If Not m_BTimed Then
          Exit For
        End If
      Next J
      DrawPoint J2, I, EEnd
      If Not m_BTimed Then
        Exit For
      End If
    Next I
   Case eRoughV
    DrawPoint 1, 1, eStart
    CurrentWidth = 10
    For J = 1 To J2 Step 10
      DrawPoint J, 0, eStart
      For I = 1 To I2 Step 10
        CurrentWidth = Rnd * 5 + 10
        DrawPoint J + Rnd * 5 - 10, I + Rnd * 5 - 10, eDraw
        DoEvents
        If Not m_BTimed Then
          Exit For
        End If
      Next I
      DrawPoint J, I2, EEnd
      If Not m_BTimed Then
        Exit For
      End If
    Next J
   Case ePointDot
    DrawPoint 1, 1, eStart
    InitWid = CurrentWidth
    For I = 1 To I2 Step InitWid
      For J = 1 To J2 Step InitWid
        CurrentWidth = Rnd * InitWid + InitWid / 2
        RI = Rnd * I2
        Rj = Rnd * J2
        DrawPoint Rj, RI, eStart
        DrawPoint Rj, RI, EEnd
        DoEvents
        If Not m_BTimed Then
          Exit For
        End If
      Next J
      If Not m_BTimed Then
        Exit For
      End If
    Next I
    Case ePointRain
    DrawPoint 1, 1, eStart
    InitWid = CurrentWidth
    For I = 1 To I2 Step InitWid
      For J = 1 To J2 Step InitWid
        CurrentWidth = Rnd * InitWid + InitWid / 2
        RI = Rnd * I2
        Rj = Rnd * J2
        RI2 = RI + 10 - Int(Rnd * 5)
        Rj2 = Rj + 10 - Int(Rnd * 5)
        DrawPoint Rj, RI, eStart
        DrawPoint Rj2, RI2, EEnd
        DoEvents
        If Not m_BTimed Then
          Exit For
        End If
      Next J
      If Not m_BTimed Then
        Exit For
      End If
    Next I
   Case ePointDash
    DrawPoint 1, 1, eStart
    InitWid = CurrentWidth
    For I = 1 To I2 Step InitWid
      For J = 1 To J2 Step InitWid
        CurrentWidth = Rnd * InitWid + InitWid / 2
        RI = Rnd * I2
        Rj = Rnd * J2
        RI2 = RI + CurrentWidth - Int(Rnd * CurrentWidth * 2)
        Rj2 = Rj + CurrentWidth - Int(Rnd * CurrentWidth * 2)
        DrawPoint Rj, RI, eStart
        DrawPoint Rj2, RI2, EEnd
        DoEvents
        If Not m_BTimed Then
          Exit For
        End If
      Next J
      If Not m_BTimed Then
        Exit For
      End If
    Next I
   Case ePointInteract
    DrawPoint 1, 1, eStart
    InitWid = CurrentWidth
    For I = 1 To I2 Step InitWid
      For J = 1 To J2 Step InitWid
        RI = Rnd * I2
        Rj = Rnd * J2
        RI2 = RI + 10 - Int(Rnd * 20)
        Rj2 = Rj + 10 - Int(Rnd * 20)
        DrawPoint Rj, RI, eStart
        DrawPoint Rj2, RI2, EEnd
        DoEvents
        If Not m_BTimed Then
          Exit For
        End If
      Next J
      If Not m_BTimed Then
        Exit For
      End If
    Next I
  End Select
  m_Canvas.MousePointer = vbDefault

End Sub

Private Property Get Canvas() As Form

  Set Canvas = m_Canvas

End Property

Private Property Set Canvas(ByVal Frm As Form)

  If Not bAbort Then
    If hasAutoRedraw(Frm) Then
      Set m_Canvas = Frm
      strBaseCaption = m_Canvas.Caption
      m_Canvas.AutoRedraw = True
      CurrentWidth = 1
      ReDim Sketch(1000) As DrawData
     Else
      MsgBox "Control " & Frm.Name & " cannot be used as a Canvas by clsSketcher"
    End If
  End If

End Property

Private Sub Class_Initialize()

  m_Scale = 1
  Speed = 20
  CurrentColour = 0

End Sub

Private Sub Class_Terminate()

  bAbort = True
  If bDirty Then
    If vbYes = MsgBox("Save before closing?", vbYesNo) Then
      SaveSket
    End If
  End If
  Do While bSaving
'suspend Close until Save is complete
  Loop

End Sub

Public Sub Cls()

  m_Canvas.Cls

End Sub

Public Property Get CurrentColour() As Long

  CurrentColour = m_Colour

End Property

Public Property Let CurrentColour(ByVal col As Long)

  m_Colour = col

End Property

Public Property Get CurrentScale() As Single

  CurrentScale = m_Scale

End Property

Public Property Let CurrentScale(ByVal scl As Single)

  m_Scale = scl

End Property

Public Property Get CurrentWidth() As Long

  CurrentWidth = m_Wid

End Property

Public Property Let CurrentWidth(wid As Long)

  If wid < 1 Then
    wid = 1
  End If
  m_Wid = wid
  m_Canvas.DrawWidth = m_Wid

End Property

Public Sub DoColorDlg()

  If bFileAccess Then
    With m_Cdlg
      .Flags = cdlCCRGBInit Or cdlCCFullOpen
      .Color = CurrentColour
      .ShowColor
      CurrentColour = .Color
    End With
   Else
    NoFileAccess
  End If

End Sub

Private Sub DoLoadSket()

  Dim arrLoad  As Variant
  Dim CTable   As Boolean
  Dim I        As Long
  Dim F        As Long
  Dim arrtmp   As Variant
  Dim strdata  As String
  Dim tmpASize As Long

  F = FreeFile
  ReDim Sketch(1000) As DrawData
  Open SketchFile For Input As #F
  Do While Not EOF(F)
    Line Input #F, strdata
    If Not CTable Then
      arrLoad = Split(strdata, ",")
      ReDim arrColourTable(UBound(arrLoad)) As Long
      For I = LBound(arrColourTable) To UBound(arrColourTable)
        arrColourTable(I) = CLng("&H" & (arrLoad(I)))
      Next I
      CTable = True
     Else
      arrtmp = Split(strdata, ",")
      With Sketch(tmpASize)
        .SX = CLng("&H" & arrtmp(0))
        .SY = CLng("&H" & arrtmp(1))
        .SMode = arrtmp(2)
        .sCol = arrColourTable(CLng("&H" & arrtmp(3)))
        .sWid = CLng("&H" & arrtmp(4))
      End With
      tmpASize = tmpASize + 1
      If tmpASize = UBound(Sketch) Then
        ReDim Preserve Sketch(UBound(Sketch) + 10000) As DrawData
      End If
    End If
  Loop
  Close #F
  ReDim Preserve Sketch(tmpASize) As DrawData
  DrawingPoint = tmpASize
  bDirty = False

End Sub

Private Sub DoSave(strF As String)

  Dim F                 As Long
  Dim I                 As Long
  Dim strSaveThis       As String
  Dim strDuplicateGuard As String
  Dim strDuplicateGuardtest As String
Dim Perc As Long
  If bDirty Then
  m_Canvas.Caption = "Saving Please Wait"
    bSaving = True
    F = FreeFile
    If LenB(strF) = 0 Then
      SketchFile = strDefaultFile
     Else
      SketchFile = strF
    End If
    m_Canvas.MousePointer = vbHourglass
    ReDim arrColourTable(0) As Variant
    arrColourTable(0) = 0
    For I = LBound(Sketch) To UBound(Sketch)
    If I Mod 100 = 0 Then
    m_Canvas.Caption = "Step 1 Getting ColourTable ...." & CLng(I / UBound(Sketch) * 100) & "%"
    End If
      fillColArray Sketch(I).sCol
    DoEvents
    Next I
   
       For I = LBound(arrColourTable) To UBound(arrColourTable)
    If I Mod 100 = 0 Then
    m_Canvas.Caption = "Step 2 Compacting ColourTable ...." & CLng(I / UBound(arrColourTable) * 100) & "%"
    End If

      arrColourTable(I) = Hex$(arrColourTable(I))
      DoEvents
    Next I
    arrColourTable = QuickSortArray(arrColourTable)
   
    Open SketchFile For Output As #F
    Print #F, Join(arrColourTable, ",")
    For I = LBound(Sketch) To UBound(Sketch)
    If I Mod 100 = 0 Then
    m_Canvas.Caption = "Step 3 Saving ...." & CLng(I / UBound(Sketch) * 100) & "%"
    End If
      DoEvents
' very simple and not very good) compression trick
    strDuplicateGuardtest = Hex$(Sketch(I).SX) & "," & Hex$(Sketch(I).SY) & "," & Sketch(I).SMode & "," & Hex$(PosInQSortArray(arrColourTable, Hex$(Sketch(I).sCol))) & "," & Hex$(Sketch(I).sWid)
    If strDuplicateGuardtest <> strDuplicateGuard Then
      strSaveThis = strSaveThis & IIf(LenB(strSaveThis), vbNewLine, "") & strDuplicateGuardtest
        strDuplicateGuard = strDuplicateGuardtest
     End If
     If I Mod 1000 = 0 Then
        Print #F, strSaveThis
        strSaveThis = ""
      End If
      
    Next I
    If LenB(strSaveThis) Then
    Print #F, strSaveThis
    End If
    strSaveThis = ""
    Close #F
    ShowCaptionHelp
    m_Canvas.MousePointer = vbDefault
    bSaving = False
    bDirty = False
  End If

End Sub
Public Function QuickSortArray(ByVal A As Variant, _
                               Optional Ascending As Boolean = True) As Variant
  On Error GoTo Not_AnArray
  If IsEmpty(A) Then
    QuickSortArray = Split("")
   Else
    QuickSortAD A, LBound(A), UBound(A), Ascending
    QuickSortArray = A
  End If
Exit Function
Not_AnArray:
  QuickSortArray = Split("")
End Function
Private Sub QuickSortAD(AnArray As Variant, _
                        Lo As Long, _
                        Hi As Long, _
                        Optional Ascending As Boolean = True)
  Dim NewHi      As Long
  Dim CurElement As Variant
  Dim NewLo      As Long
  NewLo = Lo
  NewHi = Hi
  CurElement = AnArray((Lo + Hi) / 2)
  Do While (NewLo <= NewHi)
    If Ascending Then
      Do While AnArray(NewLo) < CurElement And NewLo < Hi 'Ascending Core
        NewLo = NewLo + 1
      Loop
      Do While CurElement < AnArray(NewHi) And NewHi > Lo
        NewHi = NewHi - 1
      Loop
     Else
      Do While AnArray(NewLo) > CurElement And NewLo < Hi 'Descending Core
        NewLo = NewLo + 1
      Loop
      Do While CurElement > AnArray(NewHi) And NewHi > Lo
        NewHi = NewHi - 1
      Loop
    End If
    If NewLo <= NewHi Then
      SwapAnyThing AnArray(NewLo), AnArray(NewHi)
      NewLo = NewLo + 1
      NewHi = NewHi - 1
    End If
  Loop
  If Lo < NewHi Then
    QuickSortAD AnArray, Lo, NewHi, Ascending
  End If
  If NewLo < Hi Then
    QuickSortAD AnArray, NewLo, Hi, Ascending
  End If
End Sub
Private Sub SwapAnyThing(Var1 As Variant, _
                         Var2 As Variant)
  Dim Var3 As Variant
  Var3 = Var1
  Var1 = Var2
  Var2 = Var3
End Sub
Public Function PosInQSortArray(ByVal SortedArray As Variant, _
                             ByVal FindMe As String) As Long
  Dim Low    As Long
  Dim Middle As Long
  Dim High   As Long
  Dim Trap   As Boolean
  Dim TestMe As Variant
  If Not IsEmpty(SortedArray) Then
    If Not IsMissing(SortedArray) Then
      If UBound(SortedArray) > -1 Then
'Binary search module very fast but requires array to be sorted
        Low = LBound(SortedArray)
        High = UBound(SortedArray)
        If High >= Low Then
' invert for Descending sorted Arrays
          If SortedArray(Low) > SortedArray(High) Then
            SwapAnyThing Low, High
          End If
          High = High + 1
          Do Until High - Low = 0
            Middle = (Low + High) \ 2
' see note below*
            If Trap Then
              Middle = Low
              High = Low
            End If
            TestMe = SortedArray(Middle) ' assign once to test twice
            If TestMe >= FindMe Then
' Only tests half the time
              If TestMe = FindMe Then
                PosInQSortArray = Middle
                Exit Do 'Function
              End If
              High = Middle
             Else
              Low = Middle
            End If
            Trap = (Low = High - 1)
          Loop
         ElseIf High = Low Then
'single member test
        If SortedArray(Low) = FindMe Then
          PosInQSortArray = Low
          End If
        End If
      End If
    End If
  End If
End Function
Private Sub DrawOne(ByVal Apoint As Long)

  Select Case Sketch(Apoint).SMode
   Case eStart
    m_Canvas.PSet ((Sketch(Apoint).SX + OffsetX) * m_Scale, (Sketch(Apoint).SY + OffsetY) * m_Scale), Sketch(Apoint).sCol
   Case eDraw, EEnd
    If Apoint > 0 Then
      m_Canvas.Line ((Sketch(Apoint - 1).SX + OffsetX) * m_Scale, (Sketch(Apoint - 1).SY + OffsetY) * m_Scale)-((Sketch(Apoint).SX + OffsetX) * m_Scale, (Sketch(Apoint).SY + OffsetY) * m_Scale), Sketch(Apoint).sCol
    End If
   Case EAuto
    m_Canvas.Line ((Sketch(Apoint).SX + OffsetX) * m_Scale, (Sketch(Apoint).SY + OffsetY) * m_Scale)-((StartX + OffsetX) * m_Scale, (StartY + OffsetY) * m_Scale), Sketch(Apoint).sCol
    Sketch(Apoint).SMode = EEnd
  End Select

End Sub

Public Sub DrawPoint(ByVal X As Single, _
                     ByVal Y As Single, _
                     ByVal strMode As eDmode)

  If Not bWorking Then
    If UBound(Sketch) < DrawingPoint Then
      ReDim Preserve Sketch(UBound(Sketch) + 10000) As DrawData
    End If
    If m_PenMode = eDuplicate Then
      CurrentColour = m_Canvas.Point(X, Y)
    End If
    With Sketch(DrawingPoint)
      .SX = X
      .SY = Y
      .SMode = strMode
      .sCol = CurrentColour
      .sWid = CurrentWidth
    End With
    If strMode = EEnd Then
      ReDim Preserve Sketch(DrawingPoint) As DrawData
    End If
    If Sketch(DrawingPoint).SMode = eStart Then
      StartX = X
      StartY = Y
    End If
    DrawOne DrawingPoint
    DrawingPoint = DrawingPoint + 1
    bDirty = True
  End If

End Sub

Private Sub fillColArray(lngCol As Long)
  If PosInArray(arrColourTable, lngCol) = -1 Then 'Else
    ReDim Preserve arrColourTable(UBound(arrColourTable) + 1) As Variant
    arrColourTable(UBound(arrColourTable)) = lngCol
  End If

End Sub

Public Function hasAutoRedraw(c As Object) As Boolean

' test whether a control has autoRedraw by trying to set it

  On Error GoTo noAutoReDraw ' if error generated then skip line that sets True
  c.AutoRedraw = True
  hasAutoRedraw = True
noAutoReDraw:

End Function

Public Property Get HasData() As Boolean

'check that there is data used to activate Save menu in Demo

  If UBound(Sketch) Then
    HasData = Sketch(0).SMode + Sketch(1).SMode > 0 'no data:no save
  End If

End Property

Public Sub Initialize(Frm As Form, _
                      Optional cdl As CommonDialog)

  Set Canvas = Frm
  m_Canvas.ScaleMode = vbPixels
  m_Canvas.KeyPreview = True
  If Not cdl Is Nothing Then
    Set m_Cdlg = cdl
    bFileAccess = True
  End If

End Sub

Public Sub LoadSket(Optional BTimed As Boolean = False)

  If bFileAccess Then
    With m_Cdlg
      .Flags = cdlOFNOverwritePrompt
      .Filter = "Sketcher Files|*.Sket"
      .ShowOpen
      If Len(.FileName) Then
        SketchFile = .FileName
        DoLoadSket
        RedrawEngine BTimed
      End If
    End With
   Else
    NoFileAccess
  End If

End Sub

Public Sub LoadUnderLay()

  If bFileAccess Then
    With m_Cdlg
      .Filter = "Graphic Files|*.bmp;*.jpg;*.jpeg;*.gif|BMP files|*.bmp|JPG files|*.jpg;*.jpeg|GIF FIles|*.gif"
      .FilterIndex = 1
      .FileName = ""
      .ShowOpen
      If LenB(.FileName) Then
        strUnderlayFile = .FileName
        m_Canvas.Picture = LoadPicture(strUnderlayFile)
      End If
    End With
    m_UsingUnderLay = True
    bUnderLayToggleVisible = True
   Else
    NoFileAccess
  End If

End Sub

Private Sub m_Canvas_KeyDown(KeyCode As Integer, _
                             Shift As Integer)

'HotKeys

  Select Case KeyCode
   Case vbKeyEscape ' stop timed drawing
    m_BTimed = False
   Case vbKeyF2      'drawing width
    If Shift Then
      If CurrentWidth >= 2 Then
        CurrentWidth = CurrentWidth - 1
      End If
     Else
      CurrentWidth = CurrentWidth + 1
    End If
   Case vbKeyF3       ' Pen toggle
    PenMode = IIf(PenMode = eStandard, eDuplicate, eStandard)
   Case vbKeyF4      'Speed
    If Shift Then
      If Speed >= 2 Then
        Speed = Speed - 1
      End If
     Else
      Speed = Speed + 1
    End If
   Case vbKeyF5 '    Draw
    RedrawEngine True, Shift = 0
   Case vbKeyF6     'Scale
    If Shift Then
      If m_Scale > 0.1 Then
        m_Scale = m_Scale - 0.1
      End If
     Else
      m_Scale = m_Scale + 0.1
    End If
   Case vbKeyF7
    Auto_Sketch eRandom
   Case vbKeyBack ' undo
    Undo1
  End Select
  ShowCaptionHelp

End Sub

Private Sub m_Canvas_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

  If Button = 1 Or Button = 2 Then
    DrawPoint X, Y, eStart
  End If
  If Button = 4 Then
    CurrentColour = m_Canvas.Point(X, Y)
  End If

End Sub

Private Sub m_Canvas_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

  ShowCaptionHelp
  If Button = 1 Or Button = 2 Then
    DrawPoint X, Y, eDraw
  End If 'End If

End Sub

Private Sub m_Canvas_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

  If Button = 1 Then
    DrawPoint X, Y, EEnd
   ElseIf Button = 2 Then
    DrawPoint X, Y, EAuto
  End If 'End If

End Sub

Private Sub m_Canvas_Unload(Cancel As Integer)

  Class_Terminate

End Sub

Public Sub NewPic()

  SketchFile = strDefaultFile
  DrawingPoint = 0
  ReDim Sketch(1000) As DrawData
  m_Canvas.Cls

End Sub

Private Sub NoFileAccess()

  MsgBox "You need to initialise 'ClsSketcher' with a CommonDialog to use this routine"

End Sub

Private Sub PauseSeconds(ByVal nSecond As Single)

  Dim t0 As Single

  t0 = Timer
  Do While Timer - t0 < nSecond
    DoEvents
' if we cross midnight, back up one day
    If Timer < t0 Then
      t0 = t0 - 24 * 60 * 60
    End If
  Loop

End Sub

Public Property Get PenMode() As ePen

  PenMode = m_PenMode

End Property

Public Property Let PenMode(md As ePen)

  m_PenMode = md

End Property

Private Function PosInArray(arr As Variant, _
                            ByVal lngFind As Long) As Long

  Dim I As Long

  PosInArray = -1
  For I = 0 To UBound(arr)
    If arr(I) = lngFind Then
      PosInArray = I
      Exit For
    End If
  Next I

End Function


Public Sub RedrawEngine(Optional ByVal BTimed As Boolean = False, _
                        Optional ByVal Direction As Boolean = True)

  Dim I       As Long
  Dim For1    As Long
  Dim For2    As Long
  Dim ForStep As Long

  If Not bWorking Then
    m_BTimed = BTimed
    bWorking = True
    m_Canvas.MousePointer = vbHourglass
    m_Canvas.Cls
    If Direction Then
      For1 = 0
      For2 = UBound(Sketch) - 1
      ForStep = 1
     Else
      For1 = UBound(Sketch) - 1
      For2 = 0
      ForStep = -1
    End If
    For I = For1 To For2 Step ForStep
      If bAbort Then
        Exit For
      End If
      If m_BTimed Then
        PauseSeconds 0.0001 * Speed
      End If
      With Sketch(I)
        m_Canvas.DrawWidth = IIf(.sWid, .sWid, 1) * m_Scale
        m_Colour = .sCol
        If .SMode = eStart Then
          StartX = .SX
          StartY = .SY
        End If
      End With 'Sketch(I)
      DrawOne I
    Next I
    bWorking = False
    m_Canvas.MousePointer = vbDefault
  End If
  m_Canvas.MousePointer = vbDefault
  m_Canvas.Refresh

End Sub

Public Sub ReLoad()

  On Error Resume Next
  If LenB(SketchFile) Then
    DoLoadSket
    RedrawEngine
  End If
  On Error GoTo 0

End Sub

Public Sub SaveBMP()

  If bFileAccess Then
    With m_Cdlg
      .Flags = cdlOFNOverwritePrompt
      If LenB(SketchFile) Then
        .FileName = Replace(SketchFile, ".Sket", ".bmp")
      End If
      .Filter = "BMP file|*.bmp"
      .ShowSave
      If Len(.FileName) Then
        SavePicture m_Canvas.Image, .FileName
      End If
    End With
   Else
    NoFileAccess
  End If

End Sub

Public Sub SaveSket(Optional ByVal bSaveAs As Boolean = False)

  Dim bBlockDefaultSave As Boolean

  If bFileAccess Then
    With m_Cdlg
      If InStr(SketchFile, strDefaultFile) Then
        bBlockDefaultSave = True
        MsgBox "It is advisable not to Save with the Default FileName '" & strDefaultFile & "'"
      End If
      .Flags = cdlOFNOverwritePrompt
      If Not bSaveAs Then
        If LenB(SketchFile) Then
          If bBlockDefaultSave Then ' if current name is Default name force user to name it
            .FileName = SketchFile
            .Filter = "Sketcher Files|*.Sket"
            .ShowSave
            If Len(.FileName) Then
              DoSave .FileName
            End If
           Else ' autosave current filename
            DoSave SketchFile
          End If
         Else
          .FileName = SketchFile
          .Filter = "Sketcher Files|*.Sket"
          .ShowSave
          If Len(.FileName) Then
            DoSave .FileName
          End If
        End If
       Else
        If LenB(SketchFile) Then
          .FileName = SketchFile
         Else
          .FileName = strDefaultFile
        End If
        .Filter = "Sketcher Files|*.Sket"
        .ShowSave
        If Len(.FileName) Then
          DoSave .FileName
        End If
      End If
    End With
   Else
    NoFileAccess
  End If

End Sub

Private Sub ShowCaptionHelp()
If Not bSaving Then
  m_Canvas.Caption = strBaseCaption & "[" & SketchFileNameOnly & "][Colour=" & CurrentColour & "][Width(F2)=" & CurrentWidth & "][Pen=" & IIf(m_PenMode = eStandard, "Standard", "Duplicate") & "][Speed(F4)=" & Speed & "][Draw(F5)][m_Scale(F6)=" & m_Scale & "]"
End If
End Sub

Public Property Get SketchFile() As String

  SketchFile = strSketchFile

End Property

Public Property Let SketchFile(ByVal strFN As String)

  strSketchFile = strFN
  SketchFileNameOnly = strSketchFile
  If InStr(SketchFileNameOnly, "\") Then
    SketchFileNameOnly = Mid$(SketchFileNameOnly, InStrRev(SketchFileNameOnly, "\") + 1)
  End If

End Property

Public Property Get Speed() As Long

  Speed = m_Speed

End Property

Public Property Let Speed(ByVal s As Long)

  m_Speed = s

End Property



Public Property Get TraceOffSetX() As eOffSet

  TraceOffSetX = OffsetX

End Property

Public Property Let TraceOffSetX(ByVal NewX As eOffSet)

  Select Case NewX
   Case eNegative
    OffsetX = -(m_Canvas.ScaleX(m_Canvas.Picture.Width) + 10)
   Case eOverLay
    OffsetX = 0
   Case ePositive
    OffsetX = m_Canvas.ScaleX(m_Canvas.Picture.Width) + 10
  End Select

End Property

Public Property Get TraceOffSetY() As eOffSet

  TraceOffSetY = OffsetY

End Property

Public Property Let TraceOffSetY(ByVal NewY As eOffSet)

  Select Case NewY
   Case eNegative
    OffsetY = -(m_Canvas.ScaleY(m_Canvas.Picture.Height) + 10)
   Case eOverLay
    OffsetY = 0
   Case ePositive
    OffsetY = m_Canvas.ScaleY(m_Canvas.Picture.Height) + 10
  End Select

End Property

Public Sub UnderLayToggle()

  If bUnderLayToggleVisible Then
    bUnderLayToggleVisible = False
    m_Canvas.Picture = LoadPicture("")
   Else
    If LenB(strUnderlayFile) Then
      bUnderLayToggleVisible = True
      m_Canvas.Picture = LoadPicture(strUnderlayFile)
     Else
      bUnderLayToggleVisible = False
    End If
  End If

End Sub

Public Property Get UnderLayVisible() As Boolean

  UnderLayVisible = bUnderLayToggleVisible

End Property

Public Sub Undo1()

  Dim I As Long

  For I = UBound(Sketch) To 0 Step -1
    If Sketch(I).SMode = eStart Then
      If I - 1 > 0 Then
        ReDim Preserve Sketch(I - 1) As DrawData
        DrawingPoint = I - 1
        bDirty = True
       Else
        NewPic
      End If
      Exit For
    End If
  Next I
  RedrawEngine

End Sub

Public Property Get UsingUnderLay() As Boolean

  UsingUnderLay = m_UsingUnderLay

End Property

':)Code Fixer V2.8.3 (6/01/2005 7:58:14 AM) 79 + 913 = 992 Lines Thanks Ulli for inspiration and lots of code.

