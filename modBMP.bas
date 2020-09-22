Attribute VB_Name = "modBMP"
'*****************************************************************
'*              Bitmap file creation subroutines                 *
'*              written by Chavdar Yordanov, 04.2001             *
'*              Email: chavo@beer.com                            *
'*              Please, don't remove this title!                 *
'*****************************************************************

Option Explicit

'-----------------------------------------------------------------
'- Creates a BMP file on the disk containing the bitmap          -
'- Please, note that bBitMap array already contains the          -
'- bitmap color information. This function just adds             -
'- the bitmap header and saves the data to the disk.             -
'-----------------------------------------------------------------
Public Sub Create24bitBitmap(ByVal bmpHeight, ByVal bmpWidth, bBitmap() As Byte, sbmpFileName As String)
    
    Dim bBytes() As Byte        'will contain a long or integer value split to bytes
    Dim bmpDiskSize As Long     'Bitmap size on the disk
    Dim bmpImgSize As Long      'Bitmap image size = height x width in pixels
    Dim cPos As Long            'Current position within the bBitMap array. Set by MergeBytes sub
    Dim i As Long, j As Long    'Counters
    Dim FNo As Integer          'The free file number
    
    Const bmpOffset = 54        'header size in bytes
    Const bmpResolution = 3780  'pels per meter default x and y resolution
    Const biSize = 40           '
    Const bitCount = 24         'color depth in bits
    Const bitType = 19778       'Letters BM as double-byte WORD (the first two bytes of the file data)
    Const bitPlanes = 1         'fixed to 1
    Const bitCompression = 0    'this bitmap is non-compressed
    
    cPos = 0
    bmpImgSize = bmpHeight * bmpWidth
    bmpDiskSize = bmpOffset + bmpImgSize * bitCount / 8
    
    SplitIntoBytes bitType, 2, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bmpDiskSize, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    cPos = cPos + 4 'skipping 2 reserved WORDs
    
    SplitIntoBytes bmpOffset, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes biSize, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bmpWidth, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bmpHeight, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bitPlanes, 2, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bitCount, 2, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bitCompression, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bmpImgSize, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bmpResolution, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    SplitIntoBytes bmpResolution, 4, bBytes()
    MergeBytes bBitmap(), bBytes(), cPos
    
    FNo = FreeFile
    Open sbmpFileName For Binary Access Write As #FNo
    Put #FNo, , bBitmap()
    Close #FNo
    
End Sub

'------------- Inserts contents of bFraction into bAll array at lCurrPosition ----
Public Sub MergeBytes(ByRef bAll() As Byte, ByRef bFraction() As Byte, ByRef lCurrPosition As Long)
    Dim i
    For i = 1 To UBound(bFraction)
        bAll(lCurrPosition + i) = bFraction(i)
    Next i
    lCurrPosition = lCurrPosition + i - 1
End Sub

