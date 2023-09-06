Attribute VB_Name = "GDI"
Option Base 1
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const CREATE_ALWAYS = 2
Public Const CREATE_NEW = 1
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CXSCREEN = 0
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYSCREEN = 1
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const TRANSPARENT = 1
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const CAPTUREBLT = &H40000000
Public Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type POLYTEXT
        X As Long
        Y As Long
        N As Long
        lpStr As String
        uiFlags As Long
        rcl As RECT
        pdx As Long
End Type
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As Long
Public Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Public Declare Function PolyBezier Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Public Declare Function PolyDraw Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
Public Declare Function PolyBezierTo Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PolylineTo Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function PolyPolygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function PolyPolyline Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
Public Declare Function PolyTextOut Lib "gdi32" Alias "PolyTextOutA" (ByVal hDC As Long, pptxt As POLYTEXT, cStrings As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociateIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function AngleArc Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Double, ByVal eSweepAngle As Double) As Long
Public Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As Long
Public Declare Function ArcTo Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As Long
Public Declare Function Chord Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const PS_SOLID = 0
Public T as Long
Public hDC as long

Private m_lPower2(0 To 31) As Long

Public Function RShift(ByVal lThis As Long, ByVal lBits As Long) As Long
   If (lBits <= 0) Then
      RShift = lThis
   ElseIf (lBits > 63) Then
      ' .. error ...
   ElseIf (lBits > 31) Then
      RShift = 0
   Else
      If (lThis And m_lPower2(31 - lBits)) = m_lPower2(31 - lBits) Then
         RShift = (lThis And (m_lPower2(31 - lBits) - 1)) * m_lPower2(lBits) Or m_lPower2(31)
      Else
         RShift = (lThis And (m_lPower2(31 - lBits) - 1)) * m_lPower2(lBits)
      End If
   End If
End Function

Public Function LShift(ByVal lThis As Long, ByVal lBits As Long) As Long
   If (lBits <= 0) Then
      LShift = lThis
   ElseIf (lBits > 63) Then
      ' ... error ...
   ElseIf (lBits > 31) Then
      LShift = 0
   Else
      If (lThis And m_lPower2(31)) = m_lPower2(31) Then
         LShift = (lThis And &H7FFFFFFF) \ m_lPower2(lBits) Or m_lPower2(31 - lBits)
      Else
         LShift = lThis \ m_lPower2(lBits)
      End If
   End If
End Function

Public Sub Init()
   m_lPower2(0) = &H1&
   m_lPower2(1) = &H2&
   m_lPower2(2) = &H4&
   m_lPower2(3) = &H8&
   m_lPower2(4) = &H10&
   m_lPower2(5) = &H20&
   m_lPower2(6) = &H40&
   m_lPower2(7) = &H80&
   m_lPower2(8) = &H100&
   m_lPower2(9) = &H200&
   m_lPower2(10) = &H400&
   m_lPower2(11) = &H800&
   m_lPower2(12) = &H1000&
   m_lPower2(13) = &H2000&
   m_lPower2(14) = &H4000&
   m_lPower2(15) = &H8000&
   m_lPower2(16) = &H10000
   m_lPower2(17) = &H20000
   m_lPower2(18) = &H40000
   m_lPower2(19) = &H80000
   m_lPower2(20) = &H100000
   m_lPower2(21) = &H200000
   m_lPower2(22) = &H400000
   m_lPower2(23) = &H800000
   m_lPower2(24) = &H1000000
   m_lPower2(25) = &H2000000
   m_lPower2(26) = &H4000000
   m_lPower2(27) = &H8000000
   m_lPower2(28) = &H10000000
   m_lPower2(29) = &H20000000
   m_lPower2(30) = &H40000000
   m_lPower2(31) = &H80000000
End Sub
Sub Evals(T)
	T=T+1
	T=T or LShift(T,8) Or LShift(T,16)+ T Or RShift(T,8) Or RShift(T,16)+LShift(T,16) Or LShift(T,8) Or T+RShift(T,16) Or RShift(T,8) Or T+T
End Sub
Sub P2()
    Randomize

    hDC = GetDC(0)
    X = GetSystemMetrics(SM_CXSCREEN)
    Y = GetSystemMetrics(SM_CYSCREEN)
    x1 = 1
    y1 = 1
    x2 = 1
    y2 = 1
	X = X - 1
	Y = Y - 1
	x1 = x1 + 1
	y1 = y1 + 1
	x2 = x2 + 1
	y2 = y2 + 1
	BitBlt hDC, Int(Fix(Rnd() * X)), Int(Fix(Rnd() * Y)), x1, y1, hDC, x1, y1, DSTINVERT
End Sub
Sub p4()
    Randomize
	Evals(T)
	T = T Mod (&H1000 - &H25) + &H25
	Beep T, 1
	X = SM_CXSCREEN
	Y = SM_CYSCREEN
	w = GetSystemMetrics(0)
	h = GetSystemMetrics(1)
	BitBlt hDC, Int(Fix(Rnd() * 75)), Int(Fix(Rnd() * 75)), w, h, hDC, Int(Fix(Rnd() * 75)), Int(Fix(Rnd() * 75)), SRCCOPY
End Sub
Sub p5()
    Randomize

    w = GetSystemMetrics(0)
    h = GetSystemMetrics(1)
	Evals(T)
	T = T Mod (&H1000 - &H25) + &H25
	Beep T, 1
    BitBlt hDC, Fix(Rnd * 2), Fix(Rnd * 2), w, h, hDC, Fix(Rnd * 2), Fix(Rnd * 2), SRCAND
End Sub
Sub p7()
    Randomize
	Evals(T)
	T = T Mod (&H1000 - &H25) + &H25
	Beep T, 1
	X = SM_CXSCREEN
	Y = SM_CYSCREEN
	w = GetSystemMetrics(0)
	h = GetSystemMetrics(1)
	BitBlt hDC, Int(Fix(Rnd() * 1000)), Int(Fix(Rnd() * 1000)), w, h, hDC, Int(Fix(Rnd() * 1000)), Int(Fix(Rnd() * 1000)), SRCCOPY
End Sub
Sub p8()
	Dim a As RECT, text As String
	Randomize
	w = GetSystemMetrics(0)
	h = GetSystemMetrics(1)

	Evals(T)
	T = T Mod (&H1000 - &H25) + &H25
	Beep T, 1
	a.left = Int(Fix(Rnd() * w))
	a.top = Int(Fix(Rnd() * h))
	a.Right = Int(Fix(Rnd() * w))
	a.Bottom = Int(Fix(Rnd() * h))
	'SetBkColor hdc, TRANSPARENT
	text = "Your computer is destroyed"
	DrawText hDC, text, Len(text), a, DT_CENTER
	TextOut hDC, Int(Fix(Rnd() * w)), Int(Fix(Rnd() * h)), text, Len(text)
End Sub

Sub p9()
	Randomize

	w = GetSystemMetrics(0)
	h = GetSystemMetrics(1)
	w = w + 1
	h = h + 1
	Evals(T)
	T = T Mod (&H1000 - &H25) + &H25
	Beep T, 1
	BitBlt hDC, Int(Fix(Rnd() * w)), Int(Fix(Rnd() * h)), Int(Fix(Rnd() * w)), Int(Fix(Rnd() * h)), hDC, Int(Fix(Rnd() * w)), Int(Fix(Rnd() * h)), SRCCOPY
	BitBlt hDC, w, h, GetSystemMetrics(0), GetSystemMetrics(1), hDC, w - GetSystemMetrics(0), h - GetSystemMetrics(1), CAPTUREBLT
	ReleaseDC GetDesktopWindow(), hDC
End Sub
