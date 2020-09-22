Attribute VB_Name = "ScreenAPI"
'----- I downloaded this module from the Web some time ago --------
'-      but can't remember where exactly from :-(                 -
'-      Anyway, it does the screen capturing job so we can        -
'-      take a snapshot into frmScreen                            -
'-      and pick the color from there                             -
'------------------------------------------------------------------

Option Explicit
Option Base 0
Private Type PALETTEENTRY
         peRed As Byte
         peGreen As Byte
         peBlue As Byte
         peFlags As Byte
End Type
Private Type LOGPALETTE
         palVersion As Integer
         palNumEntries As Integer
         palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors
End Type
Private Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(7) As Byte
End Type
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
            ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
            ByVal hDC As Long, ByVal nWidth As Long, _
            ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" ( _
            ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" ( _
            ByVal hDC As Long, ByVal wStartIndex As Long, _
            ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
            As Long
Private Declare Function CreatePalette Lib "GDI32" ( _
            lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" ( _
            ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" ( _
            ByVal hDCDest As Long, ByVal XDest As Long, _
            ByVal YDest As Long, ByVal nWidth As Long, _
            ByVal nHeight As Long, ByVal hDCSrc As Long, _
            ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
            As Long
Private Declare Function DeleteDC Lib "GDI32" ( _
            ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () _
            As Long
Private Declare Function SelectPalette Lib "GDI32" ( _
            ByVal hDC As Long, ByVal hPalette As Long, _
            ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" ( _
            ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" ( _
            ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32" ( _
            ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" ( _
           ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" ( _
            ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
Private Declare Function OleCreatePictureIndirect _
            Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
            ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
            
Public Function CaptureWindow(ByVal hWndSrc As Long, _
            ByVal Client As Boolean, ByVal LeftSrc As Long, _
            ByVal TopSrc As Long, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As Picture
            Dim hDCMemory As Long
            Dim hBmp As Long
            Dim hBmpPrev As Long
            Dim R As Long
            Dim hDCSrc As Long
            Dim hPal As Long
            Dim hPalPrev As Long
            Dim RasterCapsScrn As Long
            Dim HasPaletteScrn As Long
            Dim PaletteSizeScrn As Long
        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                      ' window
        Dim LogPal As LOGPALETTE
         ' Depending on the value of Client get the proper device context
        If Client Then
            hDCSrc = GetDC(hWndSrc) ' Get device context for client area
        Else
            hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                          ' window
        End If
         ' Create a memory device context for the copy process
         hDCMemory = CreateCompatibleDC(hDCSrc)
         ' Create a bitmap and place it in the memory DC
         hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
         hBmpPrev = SelectObject(hDCMemory, hBmp)
         ' Get screen properties
         RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                            'capabilities
         HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                              'support
         PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                              ' palette
         ' If the screen has a palette make a copy and realize it
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            ' Create a copy of the system palette
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            R = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            ' Select the new palette into the memory DC and realize it
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            R = RealizePalette(hDCMemory)
        End If
         ' Copy the on-screen image into the memory DC
         R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
            LeftSrc, TopSrc, vbSrcCopy)
      ' Remove the new copy of the  on-screen image
         hBmp = SelectObject(hDCMemory, hBmpPrev)
         ' If the screen has a palette get back the palette that was
         ' selected in previously
        If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
        End If
         ' Release the device context resources back to the system
         R = DeleteDC(hDCMemory)
         R = ReleaseDC(hWndSrc, hDCSrc)
         ' Call CreateBitmapPicture to create a picture object from the
         ' bitmap and palette handles.  Then return the resulting picture
         ' object.
         Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    
End Function
     
Public Function CaptureScreen() As Picture
      
    Dim hWndScreen As Long
      
         ' Get a handle to the desktop window
        hWndScreen = GetDesktopWindow()
         ' Call CaptureWindow to capture the entire desktop give the handle
         ' and return the resulting Picture object
        Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
            Screen.Width \ Screen.TwipsPerPixelX, _
            Screen.Height \ Screen.TwipsPerPixelY)
     End Function
 
Public Function CreateBitmapPicture(ByVal hBmp As Long, _
        ByVal hPal As Long) As Picture

    Dim R As Long
    Dim Pic As PicBmp
' IPicture requires a reference to "Standard OLE Types"
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
' Fill in with IDispatch Interface ID
    With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
    End With
' Fill Pic with necessary parts
         With Pic
            .Size = Len(Pic)          ' Length of structure
            .Type = vbPicTypeBitmap   ' Type of Picture (bitmap)
            .hBmp = hBmp              ' Handle to bitmap
            .hPal = hPal              ' Handle to palette (may be null)
         End With         ' Create Picture object
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
' Return the new Picture object
    Set CreateBitmapPicture = IPic
End Function



