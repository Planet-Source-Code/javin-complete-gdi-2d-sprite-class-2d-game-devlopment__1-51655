VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sprite2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class was designed to simplify "sprite" based applications
'and even to just simplify any bitmap manipulation.
'Big props to Smidge204 from #VB on DalNET for pointing me in the direction of the
'DIBSection APIs.  A lot of speed and functionality gains came from that.
'If you find these functions handy, shoot me an E-Mail at:  Javin@Javin-Inc.com
Option Explicit

'Always alphabetize your declarations.  Keeps you from duplicating declarations.
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc&, ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Const IM_DEFAULTCOLOR = &H0 'Load with default colors.  (Most of this is here for posterity.)
Private Const IM_LOADFROMFILE = &H10 'Load from file
Private Const IM_DEFAULTSIZE = &H40 'Load with normal size.
Private Const IMAGE_BITMAP = 0 'Specifies image is a bitmap.
Private Const HIMETRIC_OFFSET = 26.4583395804413  'Necessary for converting an IPicureDisp.Width to pixels.
Private Const NULL_BRUSH = 5 'To be used for the "transparent" rectangle (the fill brush is null)
Private Const NULL_PEN = 8 'Same as the brush, but for pens.  Go figure.

'These are the different pen styles:
Public Enum PenStyleType
    PS_SOLID = 0                   'Solid line
    PS_DASH = 1                    '  -------
    PS_DOT = 2                     '  .......
    PS_DASHDOT = 3                 '  _._._._
    PS_DASHDOTDOT = 4              '  _.._.._
End Enum
Private lngPenStyle As PenStyleType, lngPenWidth As Long, lngPenColor As Long

'And now the brush color:
Private lngBrushColor As Long

'All the fun stuff necessary for GDI Bitmap Sections.
Private Type BITMAPFILEHEADER
 bfType As Integer
 bfSize As Long
 bfReserved1 As Integer
 bfReserved2 As Integer
 bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
 biSize As Long
 biWidth As Long
 biHeight As Long
 biPlanes As Integer
 biBitCount As Integer
 biCompression As Long
 biSizeImage As Long
 biXPelsPerMeter As Long
 biYPelsPerMeter As Long
 biClrUsed As Long
 biClrImportant As Long
End Type

Private Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
 bmiColors As RGBQUAD
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

'The address of the Device Context that will be holding the main sprite.
Private SpriteHdc  As Long

'This will hold the information about the main image. (Remember that each sprite could
'ALSO have a background, and mask image associated with it.)
Private BytesPerScanline As Long, BitmapPointerInfo As Long

'And the Device Context of the Mask, if they use one.
Private MaskHdc As Long

'The address of the "old" bitmap that will have to be replaced before destruction.
Private OldSpriteBitmap As Long, SpriteArray As SAFEARRAY2D

'The address of the "old" Mask bitmap, assuming they use a mask. (gotta keep track of the
'old stuff to destroy the Device Contexts properly.)
Private OldMaskBitmap As Long

'The "BackBuffer" will be used for the "Erase" procedure.  It will store the image
'that was in the position of the sprite prior to drawing the sprite there, and then
'restore that image when "Erase" is called.
Private BackBufferHdc As Long
'The address of the "old" backbuffer that will also have to be nuked before destruction.
Private OldBackBuffer As Long
'And variables that will be used by the BackBuffer
Private booBufferUsed As Boolean
Private lngBufferX As Long
Private lngBufferY As Long
Private lngBufferDC As Long

'These are self explanitory.
Private SpriteX As Long
Private SpriteY As Long
Private strTag As String
Private lngTag As Long

'These will be used if the sprite consists of numerous columns/rows.
Private CellWidth As Long
Private CellHeight As Long
Private CurrentColumn As Long
Private CurrentRow As Long

'The hDC of the 'Parent' of the sprite.  This determines the default drawing surface.
Private hdcSpriteParent As Long

'The brush and pen to be used for API drawing.
Private OldPen As Long, CurrentBrush As Long
Private OldBrush As Long, CurrentPen As Long

'To be used for displaying text on the sprite.
Private lngOldFont As Long

'Various others that don't fit in any particular category.
Private IsLoaded As Boolean
Private IsTransparent As Boolean
Private booDisabled As Boolean

Private bmpInfo As BITMAPINFO, bmpBytes() As Byte, BytesPerPixel As Long, SpriteHeight As Long, SpriteWidth As Long, HasArray As Boolean

'This is about five times faster than the LoadImageSprite function, but only works
'for .BMP bitmaps.
Public Function LoadBMPSprite(ByVal Filename As String, ByVal hdcParent As Long, Optional ByVal Columns As Long = -1, Optional ByVal Rows As Long = -1, Optional ByVal TransparencyColor As Long = -1) As Boolean
    Dim bmpInfoTemp As BITMAPINFO, lngTemp As Long
    Dim bmpTemp As Long, hdcTemp As Long, oldTemp As Long
    
    LoadBMPSprite = False
    
    'If it's already been loaded, SCREW YOU!  That's up to you to keep track of.
    If IsLoaded Then Exit Function
    
    'Create a Device Context to hold the bitmap that's compatible with the parent DC.
    hdcTemp = CreateCompatibleDC(hdcParent)
    
    'Load the bitmap image in, saving the pointer to the Bitmap info.
    bmpTemp = LoadImage(0&, Filename, IMAGE_BITMAP, 0&, 0&, IM_LOADFROMFILE + IM_DEFAULTSIZE + IM_DEFAULTCOLOR)
    
    'Get the bitmap info loaded into the bmpInfo variable.
    lngTemp = FreeFile
    Open Filename For Binary As #lngTemp
        Get #lngTemp, 15, bmpInfoTemp.bmiHeader
    Close #lngTemp

    'Move the image into the Device Context, and save the old bitmap for later.
    oldTemp = SelectObject(hdcTemp, bmpTemp)
    
    'This essentially leaves bmpTemp empty, but just for shits and grins, we'll destroy
    'anything left in it.
    DeleteObject bmpTemp
    
    With bmpInfoTemp.bmiHeader
        'Now, actually create the an empty sprite using our other function that can hold this one.
        LoadEmptySprite .biWidth, .biHeight, hdcParent, Columns, Rows
    
        'Move this sprite into it.
        BitBlt SpriteHdc, 0, 0, .biWidth, .biHeight, hdcTemp, 0, 0, vbSrcCopy
    End With
    
    'And clean this sprite up.  First, select the objects back to their original positions.
    bmpTemp = SelectObject(hdcTemp, oldTemp)
    
    'And everything gets destroyed.
    DeleteObject bmpTemp
    DeleteObject oldTemp
    DeleteDC hdcTemp
        
    IsLoaded = True
    
    'If they want it to be transparent, call the MakeTransparent function.
    If TransparencyColor <> -1 Then MakeTransparent TransparencyColor
    
    LoadBMPSprite = True
End Function


'This is the function you would use for loading .JPGs, and other non-bitmap formats.
'(Obviously, you're limited to those that are supported by the code.)  It's not nearly
'as fast as the bitmap loading, so don't use it if you're loading bitmaps.
Public Function LoadImageSprite(ByVal Filename As String, ByVal hdcParent As Long, Optional ByVal Columns As Long = -1, Optional ByVal Rows As Long = -1, Optional ByVal TransparencyColor As Long = -1) As Boolean
    Dim picTemp As IPictureDisp, hdcTemp As Long
    Dim bmpTemp As Long, bmpOldTemp As Long
    
    LoadImageSprite = False
    If IsLoaded Then Exit Function
    
    'Load the image into the IPictureDisp variable.
    Set picTemp = LoadPicture(Filename)
    
    'Grab the width/height from it (stored in himetrics.)
    bmpInfo.bmiHeader.biWidth = picTemp.Width / HIMETRIC_OFFSET
    bmpInfo.bmiHeader.biHeight = picTemp.Height / HIMETRIC_OFFSET
    
    'Create a Device Context to hold the bitmap that's compatible with the parent DC.
    hdcTemp = CreateCompatibleDC(hdcParent)
    
    'Move the image from the IPictureDisp into the Device Context.
    bmpOldTemp = SelectObject(hdcTemp, picTemp.handle)
    
    'Create our "working" device context.
    LoadEmptySprite Width, Height, hdcParent, Columns, Rows
    
    'Copy the image to our "working" DC.
    BitBlt SpriteHdc, 0, 0, Width, Height, hdcTemp, 0, 0, vbSrcCopy
    
    'Destroy the stuff you don't need anymore.
    bmpOldTemp = SelectObject(hdcTemp, bmpOldTemp)
    DeleteObject bmpOldTemp
    DeleteDC hdcTemp
    Set picTemp = Nothing
    
    IsLoaded = True
    
    If TransparencyColor <> -1 Then MakeTransparent TransparencyColor
    
    LoadImageSprite = True
End Function


'This function creates an empty device context with a bitmap in it.  Good for the
'creation of buffers and such.
Public Function LoadEmptySprite(ByVal Width As Long, ByVal Height As Long, ByVal hdcParent As Long, Optional ByVal Columns As Long = -1, Optional ByVal Rows As Long = -1, Optional ByVal UseGDI As Boolean = True, Optional ByVal BGColor As Long = 0) As Boolean
    Dim lngTemp As Long
    
    LoadEmptySprite = False
    
    'No need for cleanup, because if the sprite's already loaded, the function exits.
    If IsLoaded Then Exit Function
    
    'Store your parent information for use in EVERYTHING later.
    hdcSpriteParent = hdcParent
        
    'If column/row information has been given, then divide the sprite into column/rows.
    If Columns <> -1 Then CellWidth = Width / Columns: CurrentColumn = 1 Else CellWidth = Width
    If Rows <> -1 Then CellHeight = Height / Rows: CurrentRow = 1 Else CellHeight = Height
    
    If UseGDI Then
        
        'Store this so we know the array was used:
        HasArray = True
        
        'First, create a device context with a bitmap in it that's compatible with the "parent" DC,
        'as well as being the appropriate size.  This will also return the bytes per scanline and padding
        'byte count for use in building our array.
        PrepDIBSectionDC SpriteHdc, OldSpriteBitmap, 0&, hdcParent, Width, Height, bmpInfo, BytesPerScanline, BitmapPointerInfo

        'Now let's build the array.
        With bmpInfo.bmiHeader
                'We use the bmpBytes to access the actual pixels in RAM.
                'First create the "SafeArray2D" object.
                With SpriteArray
                    .cbElements = 1
                    .cDims = 2
                    .Bounds(0).lLbound = 0
                    ' Height of the bitmap
                    .Bounds(0).cElements = bmpInfo.bmiHeader.biHeight
                    .Bounds(1).lLbound = 0
                    ' Width of the bitmap in bits (see earlier):
                    .Bounds(1).cElements = BytesPerScanline
                    'And the pointer returned from the PrepDIBSection that points
                    'to the bitmap in the DC.
                    .pvData = BitmapPointerInfo
                End With
                'Make the bmpBytes() array point to the memory addresses, and use
                'the SAFEARRAY2D structure that we just created to build the array.
                'Since it's using this structure, we can't remove it until we're
                'done with it (which is done in the class_terminate)
                CopyMemory ByVal VarPtrArray(bmpBytes()), VarPtr(SpriteArray), 4
        
        End With

    Else
        'If we don't need per-pixel editing, don't load it as a DIB section.  Keep in mind
        'that it appears that DIB sections are the fastest option UNLESS the hDC will be
        'drawn to the screen (or a form with an auto-refresh.)  Don't ask me why, but this
        'seems to be the case.  Use DIB sections for all the stuff happening in the back-
        'ground, but use these empty bitmaps for anything rendering to screen.
        PrepBitmapDC SpriteHdc, OldSpriteBitmap, hdcParent, Width, Height
        bmpInfo.bmiHeader.biWidth = Width
        bmpInfo.bmiHeader.biHeight = Height
        'Array wasn't used.
        HasArray = False
    End If
    
    'Draw a rectangle over the whole sprite to set the background color.
    'First, create the new pen and brush, storing the old one so the SpriteHDC can later
    'be properly destroyed.
    lngTemp = CreatePen(0, 1, BGColor)
    OldPen = SelectObject(SpriteHdc, lngTemp)
    lngTemp = CreateSolidBrush(BGColor)
    OldBrush = SelectObject(SpriteHdc, lngTemp)
    'Then draw your rectangle.
    DrawRectangle 0, 0, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight
    'Note that "clearing" the background is necessary for Windows '98 and earlier machines,
    'since the memory is not Zero'd when the bitmaps are created.  In Win2K and on, the
    'memory is IN THEORY Zero'd (the back is solid black) however, for some odd reason,
    'doing this seems to give the slightest speed boost to the class.
    
    'Now create a device context to hold the background image.
    PrepBitmapDC BackBufferHdc, OldBackBuffer, hdcParent, CellWidth, CellHeight
    
    'Set the "background" mode of the DC to transparent.  This is only useful for the
    'DrawText function.
    SetBkMode SpriteHdc, 1
    
    'Empties out the GDI "error buffer" which can get too full if not called.
    GdiFlush
    
    IsLoaded = True
    
    LoadEmptySprite = True
    
End Function

'This code will create a mask for the Image2D that is passed to it, based on the
'transparency color.  It's about as fast as you're going to get for on-the-fly masking.
Public Function MakeTransparent(ByVal TransColor As Long) As Boolean
    Dim bmpTemp As Long
    Dim bmpOld As Long
    Dim lngTempColor As Long
    Dim hdcInvTemp As Long
    Dim bmpInvTemp As Long
    Dim oldInvTemp As Long
    Dim tmpMaskHDC As Long
    Dim tmpOldMaskBMP As Long
    
    MakeTransparent = False
    If Not IsLoaded Then Exit Function
    
    'If the transparency has already been set, clear the mask from RAM so you can reset
    'it.
    If IsTransparent Then
        bmpTemp = SelectObject(MaskHdc, OldMaskBitmap)
        DeleteObject bmpTemp
        DeleteObject OldMaskBitmap
        DeleteDC MaskHdc
    End If
        
    'Create a temporary DC to be used.
    tmpMaskHDC = CreateCompatibleDC(hdcSpriteParent)

    'Create a Monochrome bitmap (black and white).
    bmpTemp = CreateBitmap(bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, 1&, 1&, ByVal 0&)

    'Move the Monochrome bitmap into the Mask DC.
    tmpOldMaskBMP = SelectObject(tmpMaskHDC, bmpTemp)
    'Technically bmpTemp is now empty, since a bitmap can't be in two places at once.
    'But just to be safe, since Microsoft is full of morons, we'll clean it up anyway.
    DeleteObject bmpTemp
    
    ' Set the background color of the source DC to
    ' the color in the sprite that should be transparent.
    lngTempColor = SetBkColor(SpriteHdc, TransColor)
    
    ' Create a mask for the sprite by performing a BitBlt from
    ' the sprite's bitmap to the monochrome bitmap.
    BitBlt tmpMaskHDC, 0, 0, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, SpriteHdc, 0, 0, vbSrcCopy
    
    ' Return the background color of the DC to its normal color.
    SetBkColor SpriteHdc, lngTempColor
    
    'Create a DC to hold the inverse image of the mask (this will be used
    'to "trim off" the transparency from the original hDC).
    hdcInvTemp = CreateCompatibleDC(hdcSpriteParent)
    bmpInvTemp = CreateCompatibleBitmap(hdcSpriteParent, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight)
    oldInvTemp = SelectObject(hdcInvTemp, bmpInvTemp)
    
    'Copy an inverted image from the Mask DC to the temporary hdcInvTemp.
    BitBlt hdcInvTemp, 0, 0, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, tmpMaskHDC, 0, 0, vbNotSrcCopy
    
    'And make a regular color bitmap.
    bmpTemp = CreateCompatibleBitmap(SpriteHdc, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biWidth)
    
    'Put the colored bitmap into the Mask's DC, so you can get true white.
    bmpOld = SelectObject(tmpMaskHDC, bmpTemp)
    
    'Now we can delete bmpOld, since the ORIGINAL bmp from MaskHdc is still in
    'OldMaskBMP
    DeleteObject bmpOld
    
    'Now that we have our "inverted" mask in tmpMaskHDC, let's create a DIBSection to hold it.
    PrepDIBSectionDC MaskHdc, OldMaskBitmap, bmpOld, hdcSpriteParent, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, bmpInfo, 0&, 0&
    DeleteObject bmpOld

    'Drop the 'black' portions of the inverted mask onto the hDC to make
    'the parts that should be transparent black.
    BitBlt SpriteHdc, 0, 0, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, hdcInvTemp, 0, 0, vbSrcAnd
    
    'Get a correct copy of the mask into the new Mask DIBSection by inverting the
    'inverted mask.
    BitBlt MaskHdc, 0, 0, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, hdcInvTemp, 0, 0, vbNotSrcCopy
        
    'Put the temporary objects back where they came from.
    bmpInvTemp = SelectObject(hdcInvTemp, oldInvTemp)
    bmpOld = SelectObject(tmpMaskHDC, tmpOldMaskBMP)
    
    'Destroy all of the temporary objects.
    DeleteDC tmpMaskHDC
    DeleteDC hdcInvTemp
    DeleteObject bmpOld
    DeleteObject bmpInvTemp

    MakeTransparent = True
    IsTransparent = True
End Function

Private Sub Class_Terminate()
    Dim lngTemp As Long

    'This "releases" the RAM from the array.  If not, VB wouldn't actually release
    'the memory, even though the bitmap was "destroyed."
    CopyMemory ByVal VarPtrArray(bmpBytes), 0&, 4
    
    'I think we've beat this horse enough.  Not going to clarify what happens here.
    If lngOldFont <> 0 Then
        lngTemp = SelectObject(SpriteHdc, lngOldFont)
        DeleteObject lngTemp
    End If
    lngTemp = SelectObject(SpriteHdc, OldSpriteBitmap)
    DeleteObject lngTemp
    lngTemp = SelectObject(SpriteHdc, OldPen)
    DeleteObject lngTemp
    lngTemp = SelectObject(SpriteHdc, OldBrush)
    DeleteObject lngTemp
    DeleteObject OldSpriteBitmap
    DeleteDC SpriteHdc
    
    lngTemp = SelectObject(BackBufferHdc, OldBackBuffer)
    DeleteObject lngTemp
    DeleteObject OldBackBuffer
    DeleteDC BackBufferHdc
    
    If IsTransparent Then
        lngTemp = SelectObject(MaskHdc, OldMaskBitmap)
        DeleteObject lngTemp
        DeleteObject OldMaskBitmap
        DeleteDC MaskHdc
    End If
End Sub


Public Function DrawSprite(Optional ByVal x As Long = -999999, Optional ByVal y As Long = -999999, Optional ByVal hdcDest As Long = -1, Optional ByVal Column As Long = 0, Optional ByVal Row As Long = 0, Optional ByVal EnableErase As Boolean = True) As Boolean
    If Not IsLoaded Or booDisabled Then Exit Function
    If hdcDest = -1 Then hdcDest = hdcSpriteParent
    If x = -999999 Then x = SpriteX
    If y = -999999 Then y = SpriteY
    If Column = 0 Then Column = CurrentColumn
    If Row = 0 Then Row = CurrentRow
    
    'For starters, set the "buffer used" value to whether or not you want to be able
    'to erase the image later.  (Since that will require you to "buffer" the
    'background image.)
    booBufferUsed = EnableErase
    'If you're going to buffer, then:
    If EnableErase Then
        'Set the x,y position (obviously, the buffer is the size of one cell.)
        lngBufferX = x
        lngBufferY = y
        lngBufferDC = hdcDest
        'And grab that cell into a buffer.
        BitBlt BackBufferHdc, 0, 0, CellWidth, CellHeight, lngBufferDC, lngBufferX, lngBufferY, vbSrcCopy
    End If
    
    'Most of this is self explanitory, but for those dense ones out there, what you
    'do is (if using transparency) you use a combination of "srcand" with the mask,
    'and then "srcpaint" with the original image (where all the masked pieces have
    'been made black.)  If you aren't using transparency, a simple "srccopy" will
    'suffice.  Obviously, if not using columns and rows, you drop the whole image
    'where if you ARE using columns and rows, you would use THEM instead.
    If Column = 0 And Row = 0 Then
        If IsTransparent Then
            BitBlt hdcDest, x, y, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, MaskHdc, 0, 0, vbSrcAnd
            BitBlt hdcDest, x, y, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, SpriteHdc, 0, 0, vbSrcPaint
        Else
            BitBlt hdcDest, x, y, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, SpriteHdc, 0, 0, vbSrcCopy
        End If
    Else
        If IsTransparent Then
            BitBlt hdcDest, x, y, CellWidth, CellHeight, MaskHdc, (Column - 1) * CellWidth, (Row - 1) * CellHeight, vbSrcAnd
            BitBlt hdcDest, x, y, CellWidth, CellHeight, SpriteHdc, (Column - 1) * CellWidth, (Row - 1) * CellHeight, vbSrcPaint
        Else
            BitBlt hdcDest, x, y, CellWidth, CellHeight, SpriteHdc, (Column - 1) * CellWidth, (Row - 1) * CellHeight, vbSrcCopy
        End If
    End If
End Function

Public Function EraseSprite() As Boolean
    If Not IsLoaded Or booDisabled Or Not booBufferUsed Then Exit Function
    'Fairly simple.  To "erase" the sprite, you simply grab the image from the buffer
    'that was put there before it was drawn, and drop it on the sprite's "parent."
    BitBlt lngBufferDC, lngBufferX, lngBufferY, CellWidth, CellHeight, BackBufferHdc, 0, 0, vbSrcCopy
    booBufferUsed = False
End Function

Public Property Get hdc() As Long
    hdc = SpriteHdc
End Property

Public Property Let Parent_hDC(ByVal hdc As Long)
    hdcSpriteParent = hdc
End Property
Public Property Get Parent_hDC() As Long
    Parent_hDC = hdcSpriteParent
End Property

Public Property Get Height() As Long
    Height = bmpInfo.bmiHeader.biHeight
End Property

Public Property Get Width() As Long
    Width = bmpInfo.bmiHeader.biWidth
End Property

Public Property Get Cell_Height() As Long
    Cell_Height = CellHeight
End Property

Public Property Get Cell_Width() As Long
    Cell_Width = CellWidth
End Property

Public Property Let x(ByVal Value As Long)
    SpriteX = Value
End Property
Public Property Get x() As Long
    x = SpriteX
End Property

Public Property Let y(ByVal Value As Long)
    SpriteY = Value
End Property
Public Property Get y() As Long
    y = SpriteY
End Property

Public Property Let Column(ByVal Value As Long)
    CurrentColumn = Value
End Property
Public Property Get Column() As Long
    Column = CurrentColumn
End Property

Public Property Let Row(ByVal Value As Long)
    CurrentRow = Value
End Property
Public Property Get Row() As Long
    Row = CurrentRow
End Property

Public Property Let Tag_String(ByVal Value As String)
    strTag = Value
End Property
Public Property Get Tag_String() As String
    Tag_String = strTag
End Property

Public Property Let Tag_Long(ByVal Value As Long)
    lngTag = Value
End Property
Public Property Get Tag_Long() As Long
    Tag_Long = lngTag
End Property

Public Function SetPosAndCell(ByVal x As Long, ByVal y As Long, ByVal Column As Long, ByVal Row As Long)
    SpriteX = x
    SpriteY = y
    CurrentColumn = Column
    CurrentRow = Row
End Function

Public Function SetPosition(ByVal x As Long, ByVal y As Long)
    SpriteX = x
    SpriteY = y
End Function

Public Function SetCell(ByVal Column As Long, ByVal Row As Long)
    CurrentColumn = Column
    CurrentRow = Row
End Function

Public Property Let Enabled(ByVal Value As Boolean)
    booDisabled = Not Value
End Property
Public Property Get Enabled() As Boolean
    Enabled = Not booDisabled
End Property

Private Sub PrepDIBSectionDC(ByRef RetDC As Long, ByRef RetOldBMP As Long, ByRef RetNewBMP As Long, ByVal hdcParent As Long, ByVal Width As Long, ByVal Height As Long, ByRef bmpInfoRet As BITMAPINFO, ByRef ScanlineRet As Long, ByRef RetBitmapPointer As Long)
    Dim PadBytes As Long
    
    'First, create a device context that's compatible with the "parent" DC.
    RetDC = CreateCompatibleDC(hdcParent)
    
    With bmpInfoRet.bmiHeader
        'Store the width/height info.
        .biWidth = Width
        .biHeight = Height
    
        'Set up additional bitmap information
        .biBitCount = 24 'All bitmaps will be made 24 bit for speed and ease of access.
        .biCompression = 0&
        .biPlanes = 1
        .biSize = Len(bmpInfoRet.bmiHeader)
        
        'Go ahead and store the bytes per pixel here, so the math doesn't have
        'to be repeated in the future, as well as the height so you don't mess
        'with digging out the property.  Both of these will be used for your Set/Get
        'Pixel functions.
        BytesPerPixel = .biBitCount / 8
        SpriteHeight = .biHeight - 1
        SpriteWidth = .biWidth - 1

        'Bytes per scanline=width of the image times the bytes per pixel.
        ScanlineRet = .biWidth * BytesPerPixel

        'Padding bytes are used to keep the scanlines so the bytes across are
        'multiples of 4.
        PadBytes = 4 - (ScanlineRet Mod 4)
        If PadBytes = 4 Then PadBytes = 0
        'Now add them to the scanlines to make the scanline the appropriate length.
        ScanlineRet = PadBytes + ScanlineRet
                  
        'Fairly self explanitory.  The image size is the number of bytes across times the height.
        'Don't forget the pad bytes.
        .biSizeImage = ScanlineRet * .biHeight
         
        'Finally, create the bitmap to go into the device context.
        RetNewBMP = CreateDIBSection(hdcParent, bmpInfoRet, 0, RetBitmapPointer, 0, 0)
        
        'Place it into the DC.
        RetOldBMP = SelectObject(RetDC, RetNewBMP)
        
    End With

End Sub

Private Sub PrepBitmapDC(ByRef RetDC As Long, ByRef RetOldBMP As Long, ByVal hdcParent As Long, ByVal Width As Long, ByVal Height As Long)
    Dim bmpTemp As Long
    
    'First, create a device context that's compatible with the "parent" DC.
    RetDC = CreateCompatibleDC(hdcParent)
    
    'Most other info is useless, since these types of images can't have the bytes
    'directly accessed through the array, so there's no reason to waste the time.
        
    'Finally, create the bitmap to go into the device context.
    bmpTemp = CreateCompatibleBitmap(hdcParent, Width, Height)
    
    'Place it into the DC.
    RetOldBMP = SelectObject(RetDC, bmpTemp)
    
    'Clean up.
    DeleteObject bmpTemp

End Sub

Public Sub SetPixelRGB(ByVal x As Long, ByVal y As Long, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, Optional ByVal RelativeToCell As Boolean = False)
    If Not HasArray Then Exit Sub
    If RelativeToCell Then
        x = x + (CellHeight * (CurrentColumn - 1))
        y = y + (CellWidth * (CurrentRow - 1))
    End If
    'This is all pretty simple.  You just dump the given information into the image.
    x = x - 1: y = y - 1
    If x > SpriteWidth Or x < 0 Or y > SpriteHeight Or y < 0 Then Exit Sub
    x = x * BytesPerPixel
    y = SpriteHeight - y
    bmpBytes(x + 2, y) = Red
    bmpBytes(x + 1, y) = Green
    bmpBytes(x, y) = Blue
End Sub

Public Sub GetPixelRGB(ByVal x As Long, ByVal y As Long, ByRef Red As Long, ByRef Green As Long, ByRef Blue As Long, Optional ByVal RelativeToCell As Boolean = False)
    If Not HasArray Then Exit Sub
    If RelativeToCell Then
        x = x + (CellHeight * CurrentColumn)
        y = y + (CellWidth * CurrentRow)
    End If
    'Also pretty simple.
    x = x - 1: y = y - 1
    If x > SpriteWidth Or x < 0 Or y > SpriteHeight Or y < 0 Then Exit Sub
    x = x * BytesPerPixel
    y = SpriteHeight - y
    Red = bmpBytes(x + 2, y)
    Green = bmpBytes(x + 1, y)
    Blue = bmpBytes(x, y)
End Sub

Public Function GetPixelRed(ByVal x As Long, ByVal y As Long, Optional ByVal RelativeToCell As Boolean = False) As Long
    If Not HasArray Then Exit Function
    If RelativeToCell Then
        x = x + (CellHeight * CurrentColumn)
        y = y + (CellWidth * CurrentRow)
    End If
    'Also pretty simple.
    x = (x * BytesPerPixel) + 2
    y = SpriteHeight - y
    GetPixelRed = bmpBytes(x, y)
End Function

Public Function GetPixelGreen(ByVal x As Long, ByVal y As Long, Optional ByVal RelativeToCell As Boolean = False) As Long
    If Not HasArray Then Exit Function
    If RelativeToCell Then
        x = x + (CellHeight * CurrentColumn)
        y = y + (CellWidth * CurrentRow)
    End If
    'Also pretty simple.
    x = (x * BytesPerPixel) + 1
    y = SpriteHeight - y
    GetPixelGreen = bmpBytes(x, y)
End Function

Public Function GetPixelBlue(ByVal x As Long, ByVal y As Long, Optional ByVal RelativeToCell As Boolean = False) As Long
    If Not HasArray Then Exit Function
    If RelativeToCell Then
        x = x + (CellHeight * CurrentColumn)
        y = y + (CellWidth * CurrentRow)
    End If
    'Also pretty simple.
    x = x * BytesPerPixel
    y = SpriteHeight - y
    GetPixelBlue = bmpBytes(x, y)
End Function

Public Function GetPixelLong(ByVal x As Long, ByVal y As Long, Optional ByVal RelativeToCell As Boolean = False) As Long
    If Not HasArray Then Exit Function
    'Obviously, since these functions call the Set/GetPixelRGB functions, they
    'won't be as fast as them.  When you can, use the Set/GetPixelRGB functions.
    'I use longs here instead of bytes for a significant speed increase.
    Dim lngRed As Long, lngGreen As Long, lngBlue As Long
    GetPixelRGB x, y, lngRed, lngGreen, lngBlue, RelativeToCell
    GetPixelLong = RGB(lngRed, lngGreen, lngBlue)
End Function

Public Sub SetPixelLong(ByVal x As Long, ByVal y As Long, ByVal RGBColor As Long, Optional ByVal RelativeToCell As Boolean = False)
    If Not HasArray Then Exit Sub
    'Obviously, since these functions call the Set/GetPixelRGB functions, they
    'won't be as fast as them.  When you can, use the Set/GetPixelRGB functions
    SetPixelRGB x, y, RGBColor And &HFF, RGBColor \ 256 And &HFF, RGBColor \ 65536 And &HFF, RelativeToCell
End Sub

Public Sub SetPen(Optional ByVal Color As Long = -1, Optional ByVal Width As Long = -1, Optional ByVal Style As PenStyleType = -1)
    Dim lngTemp As Long
    
    'The following pen will be used for all GDI line drawing functions.
    If Color <> -1 Then lngPenColor = Color
    If Width <> -1 Then lngPenWidth = Width
    'Note that the dash-dot styles of pens are a little buggy.  Instead of showing
    'the background of the image between the dots, it'll show some other background
    'between the dots... Still haven't figured this one out.
    If Style <> -1 Then lngPenStyle = Style
    'Create the new "current pen."
    CurrentPen = CreatePen(lngPenStyle, lngPenWidth, lngPenColor)
    'Swap the new pen for the old pen, which is dropped into lngTemp
    lngTemp = SelectObject(SpriteHdc, CurrentPen)
    'And get rid of the old one.
    DeleteObject lngTemp
End Sub

Public Property Get PenColor() As Long
    PenColor = lngPenColor
End Property

Public Property Let PenColor(ByVal Value As Long)
    Dim lngTemp As Long
    lngPenColor = Value
    'Create the new "current pen."
    CurrentPen = CreatePen(lngPenStyle, lngPenWidth, lngPenColor)
    'Swap the new pen for the old pen, which is dropped into lngTemp
    lngTemp = SelectObject(SpriteHdc, CurrentPen)
    'And get rid of the old one.
    DeleteObject lngTemp
End Property

Public Property Get PenWidth() As Long
    PenWidth = lngPenWidth
End Property

Public Property Let PenWidth(ByVal Value As Long)
    Dim lngTemp As Long
    lngPenWidth = Value
    'Create the new "current pen."
    CurrentPen = CreatePen(lngPenStyle, lngPenWidth, lngPenColor)
    'Swap the new pen for the old pen, which is dropped into lngTemp
    lngTemp = SelectObject(SpriteHdc, CurrentPen)
    'And get rid of the old one.
    DeleteObject lngTemp
End Property

Public Property Get PenStyle() As PenStyleType
    PenStyle = lngPenStyle
End Property

Public Property Let PenStyle(ByVal Value As PenStyleType)
    Dim lngTemp As Long
    lngPenStyle = Value
    'Create the new "current pen."
    CurrentPen = CreatePen(lngPenStyle, lngPenWidth, lngPenColor)
    'Swap the new pen for the old pen, which is dropped into lngTemp
    lngTemp = SelectObject(SpriteHdc, CurrentPen)
    'And get rid of the old one.
    DeleteObject lngTemp
End Property

Public Property Get BrushColor() As Long
    BrushColor = lngBrushColor
End Property

Public Property Let BrushColor(ByVal Value As Long)
    Dim lngTemp As Long
    'The following brush will be used for all GDI fill drawing functions.
    lngBrushColor = Value
    'This is all the same as for the pen above.
    CurrentBrush = CreateSolidBrush(lngBrushColor)
    lngTemp = SelectObject(SpriteHdc, CurrentBrush)
    DeleteObject lngTemp
End Property

'This will return the Brush/Pen to its "invisible" state.
Public Sub MakeBrushInvisible()
    Dim lngTemp As Long
    CurrentBrush = GetStockObject(NULL_BRUSH)
    lngTemp = SelectObject(SpriteHdc, CurrentBrush)
    DeleteObject lngTemp
End Sub
Public Sub MakePenInvisible()
    Dim lngTemp As Long
    CurrentPen = GetStockObject(NULL_PEN)
    lngTemp = SelectObject(SpriteHdc, CurrentPen)
    DeleteObject lngTemp
End Sub

'Use the above Brush/Pen to draw a rectangle.
Public Sub DrawRectangle(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    Rectangle SpriteHdc, X1, Y1, X2, Y2
End Sub

Public Sub DrawCircle(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    'Consider this like drawing a rectangle, except that there's an ellipse inside.
    Ellipse SpriteHdc, X1, Y1, X2, Y2
End Sub

'Fills a sprite with the current brush color.
'NOTE!!!!! Because this function uses the "GetPixelLong" function, you MUST pass a
'TargetFillColor value if trying to do a fill on a sprite created with the
'"UseGDI" flag turned off.  Without GDI enabled on a sprite, the Get/Set pixel
'functions WILL NOT WORK!  The TargetFillColor should be the color of the pixels
'you intend to fill.
Public Sub Fill(ByVal x As Long, ByVal y As Long, Optional ByVal TargetFillColor = -1)
    If TargetFillColor = -1 Then TargetFillColor = GetPixelLong(x, y)
    ExtFloodFill SpriteHdc, x, y, TargetFillColor, 1 'FLOODFILLSURFACE
End Sub

'Hope I don't have to explain this one.
Public Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    MoveToEx SpriteHdc, X1, Y1, 0&
    LineTo SpriteHdc, X2, Y2
End Sub

Public Sub DrawText(ByVal Text As String, ByVal x As Long, ByVal y As Long)
    'Fairly straightforward.  Puts text on the sprite using the current font.
    TextOut SpriteHdc, x, y, Text, Len(Text)
End Sub

Public Sub ChangeFontStyle(ByVal FontName As String, ByVal lngSize As Long, Optional ByVal Rotation As Long = 0, Optional ByVal FontWeight As Long = 400, Optional ByVal Italic As Boolean = False, Optional ByVal Underline As Boolean = False, Optional ByVal Strikeout As Boolean = False)
    Dim lngTemp As Long, lngFont As Long
    lngFont = CreateFont(-MulDiv(lngSize, GetDeviceCaps(SpriteHdc, 90), 72), 0, Rotation * 10, 0, FontWeight, Italic, Underline, Strikeout, 1, 0, 0, 2, 0, FontName)
    lngTemp = SelectObject(SpriteHdc, lngFont)
    If lngOldFont = 0 Then lngOldFont = lngTemp Else DeleteObject lngTemp
End Sub

Public Sub ChangeFontColor(ByVal TextColor As Long)
    SetTextColor SpriteHdc, TextColor
End Sub
