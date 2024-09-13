Attribute VB_Name = "Module1"
Option Explicit

'Function to get screen resolution
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

'Functions to get DPI
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Const LOGPIXELSX = 88  'Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72 'A point is defined as 1/72 inches

'Return DPI
Public Function PointsPerPixel() As Double
 Dim hDC As Long
 Dim lDotsPerInch As Long

 hDC = GetDC(0)
 lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
 PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
 ReleaseDC 0, hDC
End Function


Sub FileToFolder()
    FolderSelectBox.Show
End Sub
