Attribute VB_Name = "basPainting"
Option Explicit

'Private Type BITMAP
'        bmType As Long
'        bmWidth As Long
'        bmHeight As Long
'        bmWidthBytes As Long
'        bmPlanes As Integer
'        bmBitsPixel As Integer
'        bmBits As Long
'End Type
'
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, _
        ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, _
        ByVal nWidthDest As Long, ByVal hHeightDest As Long, _
        ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
        ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
        ByVal crTransparent As Long) As Boolean
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'Public Sub DrawBitmap(ByVal hDC As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, _
'                         ByVal hBitmapSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long)
'
'    Dim hMemDc As Long, bm As BITMAP
'
'    ' 動作1:建立記憶體DC
'    hMemDc = CreateCompatibleDC(hDC)
'
'    ' 動作2: 將hBitmap選擇成為記憶體DC的Bitmap物件
'    SelectObject hMemDc, hBitmapSrc
'
'    GetObject hBitmapSrc, Len(bm), bm
'
'    ' 動作3: 將記憶體DC的圖像轉移到實際的DC
''    BitBlt hDC, nXDest, nYDest, bm.bmWidth, bm.bmHeight, hMemDc, 0, 0, vbSrcCopy
'
'    TransparentBlt hDC, nXDest, nYDest, nWidthDest, nHeightDest, _
'                    hMemDc, 0, 0, nWidthSrc, nHeightSrc, vbWhite
''    TransparentBlt hDC, nXDest, nYDest, nWidthDest, nHeightDest, _
'                    hMemDc, 0, 0, bm.bmWidth, bm.bmHeight, vbWhite
'
'    DeleteDC hMemDc
'
'End Sub
