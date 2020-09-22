Attribute VB_Name = "Mask_maker"
Public Declare Function ExtFloodFill Lib "GDI32" (ByVal hDC As Long, ByVal i As Long, ByVal i As Long, ByVal w As Long, ByVal i As Long) As Long
Public Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long

Public Const FLOODFILLSURFACE As Long = 1


Public Declare Function GetDesktopWindow Lib "User32" () As Long
Public Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

