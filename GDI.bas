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
Public hDC as long
Sub p10()
    ' Draw a Sierpinski triangle of order 5 on the picture box
    DrawSierpinski 5, Screen.width, Screen.height
End Sub

Private Sub DrawSierpinski(ByVal order As Integer, ByVal width As Long, ByVal height As Long)
    ' Get the device context of the window handle
    
    ' Create a black pen
    Dim hPen As Long
    hPen = CreatePen(PS_SOLID, 1, vbBlack)
    
    ' Select the pen into the device context
    Dim hOldPen As Long
    hOldPen = SelectObject(hDC, hPen)
    
    ' Get the width and height of the window
    Dim w As Single
    Dim h As Single
    w = width
    h = height
    
    ' Calculate the coordinates of the three vertices of the outer triangle
    Dim x1 As Single
    Dim y1 As Single
    Dim x2 As Single
    Dim y2 As Single
    Dim x3 As Single
    Dim y3 As Single
    
    x1 = w / 2 ' Top vertex
    y1 = 0
    
    x2 = 0 ' Left vertex
    y2 = h
    
    x3 = w ' Right vertex
    y3 = h
    
    ' Draw the outer triangle
    MoveToEx hDC, x1, y1, ByVal 0&
    LineTo hDC, x2, y2
    LineTo hDC, x3, y3
    LineTo hDC, x1, y1
    
    ' Draw the inner triangles recursively
    DrawTriangle hDC, order - 1, x1, y1, x2, y2, x3, y3
    
    ' Restore the original pen and delete the created pen
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    
End Sub

Private Sub DrawTriangle(ByVal hDC As Long, ByVal order As Integer, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single)
    ' Draw a triangle of the given order and coordinates
    If order > 0 Then
        T = (T + 1) Mod (&H1000 - &H25) + &H25
        Beep T, 1
        ' Calculate the midpoints of the sides
        Dim x4 As Single
        Dim y4 As Single
        Dim x5 As Single
        Dim y5 As Single
        Dim x6 As Single
        Dim y6 As Single
        
        x4 = (x1 + x2) / 2 ' Midpoint of side 1-2
        y4 = (y1 + y2) / 2
        
        x5 = (x2 + x3) / 2 ' Midpoint of side 2-3
        y5 = (y2 + y3) / 2
        
        x6 = (x3 + x1) / 2 ' Midpoint of side 3-1
        y6 = (y3 + y1) / 2
        
        ' Draw the inner triangle
        MoveToEx hDC, x4, y4, ByVal 0&
        LineTo hDC, x5, y5
        LineTo hDC, x6, y6
LineTo hDC, x4, y4
        
        ' Draw the smaller triangles recursively
        DrawTriangle hDC, order - 1, x1, y1, x4, y4, x6, y6 ' Top triangle
        DrawTriangle hDC, order - 1, x4, y4, x2, y2, x5, y5 ' Left triangle
        DrawTriangle hDC, order - 1, x6, y6, x5, y5, x3, y3 ' Right triangle
        
    End If
    
End Sub

Sub P1()
    Randomize

    X = GetSystemMetrics(SM_CXSCREEN)
    Y = GetSystemMetrics(SM_CYSCREEN)
    x1 = 1
    y1 = 1
    x2 = 1
    y2 = 1
        brush = CreateSolidBrush(RGB(Int(Fix(Rnd() * 255)), Int(Fix(Rnd() * 255)), Int(Fix(Rnd() * 255))))
        SelectObject hDC, brush
        X = X - 1
        Y = Y - 1
        x1 = x1 + 1
        y1 = y1 + 1
        x2 = x2 + 1
        y2 = y2 + 1
        PatBlt hDC, Int(Fix(Rnd() * GetSystemMetrics(SM_CXSCREEN))), Int(Fix(Rnd() * GetSystemMetrics(SM_CYSCREEN))), Int(Fix(Rnd() * GetSystemMetrics(SM_CXSCREEN))), Int(Fix(Rnd() * GetSystemMetrics(SM_CYSCREEN))), PATINVERT
        DeleteObject (brush)
        DoEvents
    ReleaseDC GetDesktopWindow(), hDC
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
Sub p3()
    Randomize

    X = GetSystemMetrics(0)
    Y = GetSystemMetrics(1)
	
        T = (T + 1) Mod (&H1000 - &H25) + &H25
        Beep T, 1
    StretchBlt hDC, -10, -10, X + 20, Y + 20, hDC, 0, 0, X, Y, SRCCOPY
    StretchBlt hDC, 10, 10, X - 20, Y - 20, hDC, 0, 0, X, Y, SRCCOPY
    ReleaseDC GetDesktopWindow(), hDC
End Sub

Sub p5()
    Randomize

    w = GetSystemMetrics(0)
    h = GetSystemMetrics(1)
	
        T = (T + 1) Mod (&H1000 - &H25) + &H25
        Beep T, 1
    BitBlt hDC, Fix(Rnd * 2), Fix(Rnd * 2), w, h, hDC, Fix(Rnd * 2), Fix(Rnd * 2), SRCAND
End Sub
Sub p6()
    Randomize

    sw = GetSystemMetrics(0)
    sh = GetSystemMetrics(1)
        StretchBlt hDC, -20, 0, sw + 40, sh, hDC, 0, 0, sw, sh, SRCCOPY
        ReleaseDC GetDesktopWindow(), hDC
End Sub
Sub p4()
    Randomize
	
        T = (T + 1) Mod (&H1000 - &H25) + &H25
        Beep T, 1
        X = SM_CXSCREEN
        Y = SM_CYSCREEN
        w = GetSystemMetrics(0)
        h = GetSystemMetrics(1)
        BitBlt hDC, Int(Fix(Rnd() * 75)), Int(Fix(Rnd() * 75)), w, h, hDC, Int(Fix(Rnd() * 75)), Int(Fix(Rnd() * 75)), SRCCOPY
End Sub

Sub p7()
    Randomize
	
        T = (T + 1) Mod (&H1000 - &H25) + &H25
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
    

        T = (T + 1) Mod (&H1000 - &H25) + &H25
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
	
        T = (T + 1) Mod (&H1000 - &H25) + &H25
        Beep T, 1
        BitBlt hDC, Int(Fix(Rnd() * w)), Int(Fix(Rnd() * h)), Int(Fix(Rnd() * w)), Int(Fix(Rnd() * h)), hDC, Int(Fix(Rnd() * w)), Int(Fix(Rnd() * h)), SRCCOPY
        BitBlt hDC, w, h, GetSystemMetrics(0), GetSystemMetrics(1), hDC, w - GetSystemMetrics(0), h - GetSystemMetrics(1), CAPTUREBLT
    ReleaseDC GetDesktopWindow(), hDC
End Sub
