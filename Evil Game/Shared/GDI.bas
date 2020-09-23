Attribute VB_Name = "modGDI"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Enum RasterOps
    blackness = &H42
    dstinvert = &H550009
    mergepaint = &HBB0226
    mergecopy = &HC000CA
    notsrccopy = &H330008
    notsrcerase = &H1100A6
    patcopy = &HF00021
    patinvert = &H5A0049
    patpaint = &HFB0A09
    srcand = &H8800C6
    srccopy = &HCC0020
    srcerase = &H440328
    srcinvert = &H660046
    srcpaint = &HEE0086
    whiteness = &HFF0062
End Enum

Enum BrushStyle
    BS_SOLID
    BS_INVISIBLE
    BS_HATCHED
End Enum

Enum HatchStyle
    HS_HORIZONTAL
    HS_VERTICAL
    HS_FDIAGONAL
    HS_BDIAGONAL
    HS_CROSS
    HS_DIAGCROSS
End Enum

Public Type LOGBRUSH
    lbStyle As Long 'Brush Style
    lbColor As Long
    lbHatch As Long 'Hatch Style
End Type


Public Const IMAGE_BITMAP As Long = 0
Public Const IMAGE_ICON As Long = 1
Public Const IMAGE_CURSOR As Long = 2

Public Const LR_LOADFROMFILE As Long = &H10
Public Const LR_LOADTRANSPARENT As Long = &H20
Public Const LR_DEFAULTSIZE As Long = &H40

Public Const GDI_ERROR As Long = &HFFFF

'Device Context
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

'Bitmap
Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function LoadImage Lib "user32.dll" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpszName As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Public Declare Function GetBitmapDimensionEx Lib "gdi32.dll" (ByVal hBitmap As Long, lpDimension As Size) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'Drawing
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32.dll" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'GDI objects
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal cbBuffer As Long, lpObject As Any) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Public Type Surface
    hDC As Long
    Width As Long
    Height As Long
End Type

Const PixelsPerHimetric As Single = 26.45833

'--------------------------------------------------------------------------------
'SURFACES
'--------------------------------------------------------------------------------

Function CreateSurface(Width As Long, Height As Long) As Surface
    'On Error Resume Next
    
    With CreateSurface
        .hDC = CreateCompatibleDC(GetDC(0))
        If .hDC Then
            Dim hBmp As Long
            
            hBmp = CreateCompatibleBitmap(GetDC(0), Width, Height)
            If hBmp Then
                If SelectObject(.hDC, hBmp) <> GDI_ERROR Then
                    .Width = Width
                    .Height = Height
                End If
                DeleteObject hBmp
            End If
        End If
    End With
End Function

Function CreateSurfaceFromFile(FileName As String, Optional Transparent As Boolean) As Surface
    'On Error Resume Next
    
    With CreateSurfaceFromFile
        .hDC = CreateCompatibleDC(GetDC(0))
        If .hDC Then
            Dim hBmp As Long
            
            hBmp = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or IIf(Transparent, LR_LOADTRANSPARENT, 0&))
            
            If hBmp Then
                If SelectObject(.hDC, hBmp) <> GDI_ERROR Then
                    Dim BITMAP As BITMAP
                    
                    If GetObject(hBmp, Len(BITMAP), BITMAP) Then
                        .Width = BITMAP.bmWidth
                        .Height = BITMAP.bmHeight
                    End If
                End If
                
                DeleteObject hBmp
            End If
        End If
    End With
End Function

Function CreateSurfaceIndirect(hDC As Long) As Surface
    'On Error Resume Next
    
    CreateSurfaceIndirect.hDC = hDC
End Function

Sub DeleteSurface(Surface As Surface)
    'On Error Resume Next
    
    With Surface
        If DeleteDC(.hDC) Then
            .hDC = 0
            .Width = 0
            .Height = 0
        End If
    End With
End Sub

'--------------------------------------------------------------------------------
'DRAWING
'--------------------------------------------------------------------------------

Sub RenderSurface(hDC As Long, Surface As Surface)
    'On Error Resume Next
    
    BitBlt hDC, 0, 0, Surface.Width, Surface.Height, Surface.hDC, 0, 0, srccopy
End Sub

Sub ClearSurface(Surface As Surface)
    'On Error Resume Next
    
    BitBlt Surface.hDC, 0, 0, Surface.Width, Surface.Height, 0, 0, 0, blackness
End Sub

Sub SetPen(Surface As Surface, Style As DrawStyleConstants, Width As Long, Color As Long)
    'On Error Resume Next
    
    Dim hPen As Long
    Dim hOldPen As Long
    
    hPen = CreatePen(Style, Width, Color)
    If hPen Then
        hOldPen = SelectObject(Surface.hDC, hPen)
        If hOldPen <> GDI_ERROR Then DeleteObject hOldPen
    End If
End Sub

Sub SetBrush(Surface As Surface, Style As BrushStyle, Optional Color As Long, Optional HatchStyle As HatchStyle)
    'On Error Resume Next
    
    Dim LOGBRUSH As LOGBRUSH
    Dim hBrush As Long
    Dim hOldBrush As Long
    
    With LOGBRUSH
        .lbStyle = Style
        .lbColor = Color
        .lbHatch = HatchStyle
    End With
    
    hBrush = CreateBrushIndirect(LOGBRUSH)
    If hBrush Then
        hOldBrush = SelectObject(Surface.hDC, hBrush)
        If hOldBrush <> GDI_ERROR Then DeleteObject hOldBrush
    End If
End Sub

Sub DrawLine(Surface As Surface, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    'On Error Resume Next
    
    Dim Point As POINTAPI
    
    MoveToEx Surface.hDC, X1, Y1, Point
    LineTo Surface.hDC, X2, Y2
End Sub

Sub DrawRectangle(Surface As Surface, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    'On Error Resume Next
    
    Rectangle Surface.hDC, X1, Y1, X2, Y2
End Sub
