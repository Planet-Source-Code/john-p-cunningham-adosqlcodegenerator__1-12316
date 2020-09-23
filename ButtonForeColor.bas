Attribute VB_Name = "ButtonForeColor"
'==================================================================
'
'   Found at Visual Basic Thunder, www.vbthunder.com
'   and modified by Ulli
'
'   This module provides an easy way to change the text color
'   of a VB CommandButton control. To use the code with a
'   CommandButton, you should:
'
'   - Set the button's Style property to "Graphical" at design time.
'
'   - Optionally set its BackColor and Picture properties.
'
'   - Call SetButtonForeColor in the Form_Load event:
'       SetButtonForeColor Command1, vbBlue, Alignment
'       (You can do this multiple times during your program's
'       execution, even without calling UnsetButtonForeColor.)
'
'   - Call UnsetButtonForeColor in the Form_Unload event:
'       UnsetButtonForeColor Command1
'
'   Unfortunately this works only for single line captions
'==================================================================
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Private Const TRANSPARENT   As Long = 1
Private Const GWL_WNDPROC   As Long = -4
Private Const ODT_BUTTON    As Long = 4
Private Const ODS_SELECTED  As Long = &H1
Private Const WM_DESTROY    As Long = &H2
Private Const WM_DRAWITEM   As Long = &H2B
Private Const DT_HCENTER    As Long = &H1
Private Const DT_TOP        As Long = &H0
Private Const DT_VCENTER    As Long = &H4
Private Const DT_BOTTOM     As Long = &H8
Private Const DT_SINGLELINE As Long = &H20

Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    ItemID      As Long
    ItemAction  As Long
    ItemState   As Long
    hWndItem    As Long
    hDC         As Long
    rcItem      As RECT
    ItemData    As Long
End Type

Public Enum AlignText
    AlignTop = DT_TOP
    AlignCenter = DT_VCENTER
    AlignBottom = DT_BOTTOM
    ThreeD = DT_VCENTER Or DT_BOTTOM
End Enum

'property names
Private Const PropCustom = "UMGCustom"
Private Const PropForeColor = "UMGForeColor"
Private Const PropAlign = "UMGVAlign"
Private Const PropSubclass = "UMGDrawProc"

Public Sub SetButtonForeColor(Button As CommandButton, ByVal ForeColor As OLE_COLOR, Optional ByVal Alignment As AlignText = AlignCenter)

  Dim hWndPnt   As Long
    
    With Button
        hWndPnt = GetParent(.hWnd)
        If GetProp(hWndPnt, PropSubclass) = 0 Then 'not yet subclassed
            SetProp hWndPnt, PropSubclass, GetWindowLong(hWndPnt, GWL_WNDPROC)
            SetWindowLong hWndPnt, GWL_WNDPROC, AddressOf DrawButtonProc
        End If
        SetProp .hWnd, PropCustom, True
        SetProp .hWnd, PropForeColor, ForeColor
        SetProp .hWnd, PropAlign, Alignment
        .Refresh
    End With

End Sub

Public Sub UnsetButtonForeColor(Button As CommandButton)

    With Button
        RemoveProp .hWnd, PropCustom
        RemoveProp .hWnd, PropForeColor
        RemoveProp .hWnd, PropAlign
        .Refresh
    End With

End Sub

Private Function DrawButtonProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim lOldProc  As Long
  Dim di        As DRAWITEMSTRUCT
  Dim s         As String
  Dim VA        As AlignText

    lOldProc = GetProp(hWnd, PropSubclass)
    DrawButtonProc = CallWindowProc(lOldProc, hWnd, wMsg, wParam, lParam)
    Select Case wMsg
      Case WM_DRAWITEM
        CopyMemory di, ByVal lParam, Len(di)
        With di
            If .CtlType = ODT_BUTTON Then
                If GetProp(.hWndItem, PropCustom) Then
                    VA = GetProp(.hWndItem, PropAlign)
                    With .rcItem
                        Select Case VA
                          Case DT_TOP
                            .Top = .Top + 4
                          Case DT_BOTTOM
                            .Bottom = .Bottom - 4
                          Case ThreeD
                            .Left = .Left - 1
                            .Top = .Top - 1
                            .Right = .Right - 1
                            .Bottom = .Bottom - 1
                            VA = AlignCenter
                        End Select
                        If (di.ItemState And ODS_SELECTED) = ODS_SELECTED Then
                            'Button is in down state - offset the text
                            .Left = .Left + 1
                            .Top = .Top + 1
                            .Right = .Right + 1
                            .Bottom = .Bottom + 1
                        End If
                    End With
                    SetBkMode .hDC, TRANSPARENT
                    s = String$(255, 0)
                    GetWindowText .hWndItem, s, Len(s)
                    s = Left$(s, InStr(s, Chr$(0)) - 1)
                    SetTextColor .hDC, GetProp(.hWndItem, PropForeColor)
                    DrawText .hDC, s, Len(s), .rcItem, DT_SINGLELINE Or DT_HCENTER Or VA
                End If
            End If
        End With
      Case WM_DESTROY
        If lOldProc Then 'is subclassed
            SetWindowLong hWnd, GWL_WNDPROC, lOldProc
            RemoveProp hWnd, PropSubclass
        End If
    End Select

End Function
