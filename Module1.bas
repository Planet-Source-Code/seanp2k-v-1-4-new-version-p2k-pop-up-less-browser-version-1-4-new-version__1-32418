Attribute VB_Name = "Colorstuff"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Type zRGB
    R As Long
    G As Long
    B As Long
End Type
Public LCT As Byte
Public LC As Boolean
Public DW As Boolean
Public tsp As String
Public hsp As String
Public HBC As String
Public Const ColorCode = "10"
Public Const ColorCode2 = "10"
Public Const ColorCode3 = "10"
Public Const ColorCode4 = "10"
Public Mode As Byte
Public hMode As Byte
Public MainMode As Byte
Public StopCode As Boolean

Public Function LongToRGB(ColorValue As Long) As zRGB
    Dim rCol As Long, gCol As Long, bCol As Long
    rCol = ColorValue And &H10000FF
    gCol = (ColorValue And &H100FF00) / (2 ^ 8)
    bCol = (ColorValue And &H1FF0000) / (2 ^ 16)
    LongToRGB.R = rCol
    LongToRGB.G = gCol
    LongToRGB.B = bCol
End Function

