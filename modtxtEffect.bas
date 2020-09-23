Attribute VB_Name = "modtxtEffect"

Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function SetTextCharacterExtra Lib "GDI32" (ByVal hDC As Long, ByVal nCharExtra As Long) As Long

Private Type RECT
    Left       As Long
    Top        As Long
    Right      As Long
    Bottom     As Long
End Type

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const COLOR_BTNFACE = 15

Private Declare Function TextOut Lib "GDI32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Private Const DT_DISPFILE = 6            '  Display-file
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_LEFT = &H0
Private Const DT_METAFILE = 5            '  Metafile, VDM
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_PLOTTER = 0             '  Vector plotter
Private Const DT_RASCAMERA = 3           '  Raster camera
Private Const DT_RASDISPLAY = 1          '  Raster display
Private Const DT_RASPRINTER = 2          '  Raster printer
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Public Sub TextEffect(obj As Object, ByVal sText As String, ByVal lX As Long, ByVal lY As Long, Optional ByVal bLoop As Boolean = False, Optional ByVal lStartSpacing As Long = 128, Optional ByVal lEndSpacing As Long = -1, Optional ByVal oColor As OLE_COLOR = vbWindowText)
   
   Dim lhDC             As Long
   Dim i                As Long
   Dim x                As Long
   Dim lLen             As Long
   Dim hBrush           As Long
   Static tR            As RECT
   Dim iDir             As Long
   Dim bNotFirstTime    As Boolean
   Dim lTime            As Long
   Dim lIter            As Long
   Dim bSlowDown        As Boolean
   Dim lCOlor           As Long
   Dim bDoIt            As Boolean
   
   lhDC = obj.hDC
   iDir = -1
   i = lStartSpacing
   tR.Left = lX: tR.Top = lY: tR.Right = lX: tR.Bottom = lY
   OleTranslateColor oColor, 0, lCOlor
   
   hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
   lLen = Len(sText)
   
   SetTextColor lhDC, lCOlor
   bDoIt = True
   
   Do While bDoIt
      lTime = timeGetTime
      If (i < -3) And Not (bLoop) And Not (bSlowDown) Then
         bSlowDown = True
         iDir = 1
         lIter = (i + 4)
      End If
      If (i > 128) Then iDir = -1
      If Not (bLoop) And iDir = 1 Then
         If (i = lEndSpacing) Then
            ' Stop
            bDoIt = False
         Else
            lIter = lIter - 1
            If (lIter <= 0) Then
               i = i + iDir
               lIter = (i + 4)
            End If
         End If
      Else
         i = i + iDir
      End If
      
      FillRect lhDC, tR, hBrush
      x = 32 - (i * lLen)
      SetTextCharacterExtra lhDC, i
      DrawText lhDC, sText, lLen, tR, DT_CALCRECT
      tR.Right = tR.Right + 4
      If (tR.Right > obj.ScaleWidth \ Screen.TwipsPerPixelX) Then tR.Right = obj.ScaleWidth \ Screen.TwipsPerPixelX
      DrawText lhDC, sText, lLen, tR, DT_LEFT
      obj.Refresh
      
      Do
         DoEvents
         If obj.Visible = False Then Exit Sub
      Loop While (timeGetTime - lTime) < 20
   
   Loop
   DeleteObject hBrush

End Sub

