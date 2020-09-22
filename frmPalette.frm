VERSION 5.00
Begin VB.Form frmPalette 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   382
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcVertical 
      Height          =   3900
      Left            =   4050
      MouseIcon       =   "frmPalette.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmPalette.frx":030A
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   1
      Top             =   75
      Width           =   330
   End
   Begin VB.PictureBox pcMain 
      ClipControls    =   0   'False
      Height          =   3900
      Left            =   60
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   60
      Width           =   3900
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pcMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lCol As Long
    lCol = pcMain.Point(X, Y)
    frmColorPicker.ShowColors lCol
End Sub

Private Sub pcVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lCol As Long
    lCol = pcVertical.Point(X, Y)
    
End Sub

Private Sub pcVertical_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMaxColor As Long
    lMaxColor = pcVertical.Point(X, Y)
    ShowPalette lMaxColor
End Sub

Public Sub ShowPalette(ByVal lMaxColor As Long)
    Dim i, j, lCol, R, G, B
    Dim R1, G1, B1, k
    B = lMaxColor \ (256 ^ 2)
    G = (lMaxColor - B * 256 ^ 2) \ 256
    R = (lMaxColor - B * 256 ^ 2 - G * 256)
    For i = 0 To 255
        For j = 0 To 255
            k = (255 - j)
            R1 = k / 255 * (i / 255 * R) + j
            G1 = k / 255 * (i / 255 * G) + j
            B1 = k / 255 * (i / 255 * B) + j
            pcMain.PSet (i, j), RGB(R1, G1, B1)
        Next j
    Next i
End Sub


