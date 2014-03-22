Attribute VB_Name = "modPicture"
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const SRCCOPY = &HCC0020

Public Function TakePicture() As String

    Dim ScreenDC As Long
    ScreenDC = GetDC(GetDesktopWindow)
    
    Dim lngLeft As Long, lngTop As Long
    lngLeft = (frmMain.Left / Screen.TwipsPerPixelX) + frmMain.picImage.Left
    lngTop = (frmMain.Top / Screen.TwipsPerPixelY) + frmMain.picImage.Top
    
    Dim lngWidth As Long, lngHeight As Long
    lngWidth = frmMain.picImage.Width
    lngHeight = frmMain.picImage.Height

    StretchBlt frmMain.picImage.hdc, 0, 0, lngWidth, lngHeight, ScreenDC, lngLeft, lngTop, lngWidth, lngHeight, SRCCOPY
    
    Dim FileName As String, FilePath As String
    FileName = Replace$(Date & ChrW$(32) & Time, "/", "-")
    FileName = Replace$(FileName, ":", "-")
    
    FilePath = App.Path & "\imgs\" & FileName & ChrW$(32) & GetTickCount & ".bmp"
    
    SavePicture frmMain.picImage.Image, FilePath
    
    TakePicture = FilePath
    
End Function
