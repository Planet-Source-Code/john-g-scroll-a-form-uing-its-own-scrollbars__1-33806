Attribute VB_Name = "ModScrollForm"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" _
        Alias "SetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)

Public Declare Function ShowScrollBar Lib "user32" _
        (ByVal hwnd As Long, ByVal wBar As Long, _
        ByVal bShow As Long) As Long
        
' Scroll bar Type constants
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_BOTH = 3

Public Declare Function SetScrollInfo Lib "user32" _
        (ByVal hwnd As Long, ByVal n As Long, _
        lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long

Public Declare Function GetScrollInfo Lib "user32" _
        (ByVal hwnd As Long, ByVal n As Long, _
        lpScrollInfo As SCROLLINFO) As Long

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

' SCROLLINFO fMask constants
Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = SIF_RANGE Or SIF_PAGE Or _
                        SIF_PAGE Or SIF_TRACKPOS

' Scroll bar message flags
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114

' Scroll Bar commands
Public Const SB_LINEUP As Long = &H0
Public Const SB_LINERIGHT As Long = &H0
Public Const SB_LINEDOWN As Long = &H1
Public Const SB_LINELEFT As Long = &H1
Public Const SB_PAGEUP As Long = &H2
Public Const SB_PAGERIGHT As Long = &H2
Public Const SB_PAGEDOWN As Long = &H3
Public Const SB_PAGELEFT As Long = &H3
Public Const SB_THUMBPOSITION As Long = &H4
Public Const SB_THUMBTRACK As Long = &H5
Public Const SB_TOP As Long = &H6
Public Const SB_RIGHT As Long = &H6
Public Const SB_BOTTOM As Long = &H7
Public Const SB_LEFT As Long = &H7
Public Const SB_ENDSCROLL As Long = &H8

Public Declare Function ScrollWindowEx Lib "user32" _
        (ByVal hwnd As Long, ByVal dx As Long, _
        ByVal dy As Long, lprcScroll As Any, _
        lprcClip As Any, ByVal hrgnUpdate As Long, _
        ByVal lprcUpdate As Any, ByVal fuScroll As Long) As Long
Public Const SW_INVALIDATE = &H2
Public Const SW_SCROLLCHILDREN = &H1

Public Declare Function UpdateWindow Lib "user32" _
        (ByVal hwnd As Long) As Long
        
Public Declare Sub CopyMemory Lib "Kernel32" _
        Alias "RtlMoveMemory" (Destination As Any, _
        Source As Any, ByVal Length As Long)

Public Declare Function IsWindow Lib "user32" _
        (ByVal hwnd As Long) As Long
        
Public Declare Function CallWindowProc Lib "user32" _
        Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
        ByVal Msg As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
        
Private Const CONST_SMALL = 5
Private sbINFOV As SCROLLINFO
Private sbINFOH As SCROLLINFO
Private m_oldWinProc As Long
Private m_hWnd As Long
Private oldvpos As Long
Private oldhpos As Long

Public Sub AddScrollBar(frm As Form)
    Dim min As Long
    Dim maxv As Long
    Dim maxh As Long
    
    frm.ScaleMode = vbPixels
    m_hWnd = frm.hwnd
    maxh = GetMaxWidth(frm)
    maxv = GetMaxHeight(frm)
    min = 0
    m_oldWinProc = SetWindowLong(m_hWnd, GWL_WNDPROC, _
            AddressOf HandleScrollMsg)
    ShowScrollBar m_hWnd, SB_BOTH, True

' Set Scroll bar info
With sbINFOV
    .cbSize = Len(sbINFOV)
    .fMask = SIF_RANGE Or SIF_PAGE
    .nMin = min
    .nMax = maxv
    .nPage = maxv \ 10
End With

With sbINFOH
    .cbSize = Len(sbINFOH)
    .fMask = SIF_RANGE Or SIF_PAGE
    .nMin = min
    .nMax = maxh
    .nPage = maxh \ 10
End With

SetScrollInfo m_hWnd, SB_VERT, sbINFOV, True
SetScrollInfo m_hWnd, SB_HORZ, sbINFOH, True
End Sub

Public Function GetMaxHeight(frm As Form) As Long
Dim yVal As Long
Dim maxVal As Long
Dim lHeight As Long
Dim ctl As Control          ' added jgd
For Each ctl In frm.Controls
    With ctl
        If Not (LCase(ctl.Name) = "picform") Then
            yVal = .Top + .Height
            If yVal > maxVal Then
                maxVal = yVal
                lHeight = .Height
            End If
        End If
    End With
Next ctl
GetMaxHeight = maxVal - lHeight
End Function

Public Function GetMaxWidth(frm As Form) As Long
Dim xVal As Long
Dim maxVal As Long
Dim lWidth As Long
Dim ctl As Control          ' added jgd
For Each ctl In frm.Controls
        If Not (LCase(ctl.Name) = "picform") Then
        With ctl
            xVal = .Left + .Width
            If xVal > maxVal Then
            maxVal = xVal
            lWidth = .Width
            End If
        End With
    End If
Next ctl
GetMaxWidth = maxVal - lWidth
End Function


Public Sub DestroyScrollBar()
    If IsWindow(m_hWnd) Then
        If m_oldWinProc Then
            SetWindowLong m_hWnd, GWL_WNDPROC, m_oldWinProc
        End If
    End If
End Sub

Public Function HandleScrollMsg(ByVal hwnd As Long, _
        ByVal uMSg As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Dim nwPos As Long
Dim ScrollUnit As Long

Select Case uMSg
    Case WM_VSCROLL
        With sbINFOV
            nwPos = .nPos
            .fMask = SIF_ALL
            GetScrollInfo m_hWnd, SB_VERT, sbINFOV
            Select Case LOWORD(wParam)
                Case SB_BOTTOM: nwPos = .nMax
                Case SB_TOP: nwPos = .nMin
                Case SB_LINEDOWN
                    If .nPos < .nMax Then _
                        nwPos = .nPos + CONST_SMALL
                Case SB_LINEUP
                    If .nPos > .nMin Then _
                        nwPos = .nPos - CONST_SMALL
                Case SB_PAGEDOWN
                    nwPos = .nPos + .nPage
                Case SB_PAGEUP: nwPos = .nPos - .nPage
                Case SB_THUMBPOSITION, SB_THUMBTRACK
                    nwPos = HIWORD(wParam)
                Case SB_ENDSCROLL: nwPos = .nPos
            End Select
            .fMask = SIF_POS
            .nPos = nwPos
        End With
        SetScrollInfo hwnd, SB_VERT, sbINFOV, True
        ScrollUnit = oldvpos - nwPos
        ScrollWindowEx m_hWnd, 0, ScrollUnit, ByVal 0&, _
        ByVal 0&, 0&, ByVal 0&, SW_SCROLLCHILDREN Or SW_INVALIDATE
        UpdateWindow m_hWnd
        HandleScrollMsg = 0&
        oldvpos = nwPos
        Exit Function
    Case WM_HSCROLL
        With sbINFOH
            nwPos = .nPos
            .fMask = SIF_ALL
            GetScrollInfo m_hWnd, SB_HORZ, sbINFOH
            Select Case LOWORD(wParam)
                Case SB_LEFT: nwPos = .nMax
                Case SB_RIGHT: nwPos = .nMin
                Case SB_LINELEFT
                    If .nPos < .nMax Then _
                        nwPos = .nPos + CONST_SMALL
                Case SB_LINERIGHT
                    If .nPos > .nMin Then _
                        nwPos = .nPos - CONST_SMALL
                Case SB_PAGELEFT: nwPos = .nPos + .nPage
                Case SB_PAGERIGHT: nwPos = .nPos - .nPage
                Case SB_THUMBPOSITION, SB_THUMBTRACK
                    nwPos = HIWORD(wParam)
                Case SB_ENDSCROLL: nwPos = .nPos
            End Select
            .fMask = SIF_POS
            .nPos = nwPos
        End With
        SetScrollInfo hwnd, SB_HORZ, sbINFOH, True
        ScrollUnit = oldhpos - nwPos
        ScrollWindowEx m_hWnd, ScrollUnit, 0, ByVal 0&, _
        ByVal 0&, 0&, ByVal 0&, SW_SCROLLCHILDREN Or SW_INVALIDATE
        UpdateWindow m_hWnd
        HandleScrollMsg = 0&
        oldhpos = nwPos
        Exit Function
    End Select
    HandleScrollMsg = CallWindowProc(m_oldWinProc, _
                    hwnd, uMSg, wParam, lParam)
                    
End Function

Public Function LOWORD(ByVal dwValue As Long) As Integer
' returns low 15-bit integer from a 32-bit long
    CopyMemory LOWORD, dwValue, 2&
End Function

Public Function HIWORD(ByVal dwValue As Long) As Integer
    ' returns high 16-bit integer from a 32-bit long
        CopyMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2&
End Function

