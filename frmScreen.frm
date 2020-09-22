VERSION 5.00
Begin VB.Form frmScreen 
   BorderStyle     =   0  'None
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2670
   Icon            =   "frmScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   90
      ScaleHeight     =   2055
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   135
      Width           =   2250
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------ picks a color from the copy of the screen in picture1 ------
Private Sub GetColorFromScreen(xMousePos As Single, yMousePos As Single)
    Dim mColor As Long
    mColor = Me.Picture1.Point(xMousePos, yMousePos)
    frmColorPicker.ShowColors mColor
End Sub

Private Sub Form_Activate()
    frmColorPicker.Show
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GetColorFromScreen x, y
    frmColorPicker.btPick.Enabled = True
    Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    GetColorFromScreen x, y
End Sub
