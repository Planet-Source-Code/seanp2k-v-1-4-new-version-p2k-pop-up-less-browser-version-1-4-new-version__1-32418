VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColorPicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Color picker"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   9045
   ClipControls    =   0   'False
   Icon            =   "frmColorPicker.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9045
   Begin MSComDlg.CommonDialog CDL 
      Left            =   2700
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.OptionButton opPal 
      Caption         =   "IExplorer 4+"
      Height          =   195
      Index           =   2
      Left            =   7020
      TabIndex        =   45
      Top             =   4185
      Width           =   1185
   End
   Begin VB.OptionButton opPal 
      Caption         =   "PC / MAC"
      Height          =   195
      Index           =   1
      Left            =   5850
      TabIndex        =   44
      Top             =   4185
      Width           =   1095
   End
   Begin VB.OptionButton opPal 
      Caption         =   "Gradient"
      Height          =   195
      Index           =   0
      Left            =   4680
      TabIndex        =   43
      Top             =   4185
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Height          =   2355
      Left            =   0
      TabIndex        =   33
      Top             =   1215
      Width           =   2220
      Begin VB.ComboBox cbWeb 
         Height          =   315
         Left            =   585
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Lists the colors supported by Internet Explorer 4+"
         Top             =   1440
         Width           =   1545
      End
      Begin VB.TextBox txCol 
         Height          =   330
         Index           =   0
         Left            =   585
         TabIndex        =   2
         Text            =   "0"
         Top             =   225
         Width           =   1095
      End
      Begin VB.TextBox txCol 
         Height          =   330
         Index           =   1
         Left            =   570
         TabIndex        =   4
         Text            =   "0"
         Top             =   630
         Width           =   1095
      End
      Begin VB.CommandButton btCopy 
         Caption         =   "&1"
         Height          =   330
         Index           =   0
         Left            =   1695
         TabIndex        =   36
         ToolTipText     =   "Copy to clipboard"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton btCopy 
         Caption         =   "&2"
         Height          =   330
         Index           =   1
         Left            =   1695
         TabIndex        =   35
         ToolTipText     =   "Copy to clipboard"
         Top             =   630
         Width           =   420
      End
      Begin VB.TextBox txCol 
         Height          =   330
         Index           =   2
         Left            =   585
         TabIndex        =   6
         Text            =   "0"
         Top             =   1035
         Width           =   1095
      End
      Begin VB.CommandButton btCopy 
         Caption         =   "&3"
         Height          =   330
         Index           =   2
         Left            =   1695
         TabIndex        =   34
         ToolTipText     =   "Copy to clipboard"
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "&IE 4+"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   1485
         Width           =   420
      End
      Begin VB.Label lbCol 
         Caption         =   "&Long:"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   600
      End
      Begin VB.Label lbCol 
         Caption         =   "&Hex:"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   675
         Width           =   600
      End
      Begin VB.Label lbCol 
         Caption         =   "&RGB:"
         Height          =   330
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   1080
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   2340
      TabIndex        =   26
      Top             =   1215
      Width           =   2220
      Begin VB.OptionButton opModify 
         Caption         =   "Web safe color"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   48
         Top             =   2070
         Width           =   1995
      End
      Begin VB.OptionButton opModify 
         Caption         =   "16-bit color"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   47
         Top             =   1845
         Width           =   1995
      End
      Begin VB.OptionButton opModify 
         Caption         =   "24-bit color"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   46
         Top             =   1620
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.HScrollBar scColor 
         Height          =   240
         Index           =   0
         LargeChange     =   16
         Left            =   630
         Max             =   255
         TabIndex        =   29
         Top             =   315
         Width           =   1500
      End
      Begin VB.HScrollBar scColor 
         Height          =   240
         Index           =   1
         LargeChange     =   16
         Left            =   630
         Max             =   255
         TabIndex        =   28
         Top             =   765
         Width           =   1500
      End
      Begin VB.HScrollBar scColor 
         Height          =   240
         Index           =   2
         LargeChange     =   16
         Left            =   630
         Max             =   255
         TabIndex        =   27
         Top             =   1215
         Width           =   1500
      End
      Begin VB.Label lbRGB 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   32
         Top             =   315
         Width           =   510
      End
      Begin VB.Label lbRGB 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   31
         Top             =   765
         Width           =   510
      End
      Begin VB.Label lbRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   30
         Top             =   1215
         Width           =   510
      End
   End
   Begin VB.CommandButton btPalette 
      Caption         =   "Pale&tte >>>"
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Top             =   3600
      Width           =   1995
   End
   Begin VB.CommandButton btExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2520
      TabIndex        =   24
      Top             =   4005
      Width           =   1995
   End
   Begin VB.CommandButton btDialog 
      Caption         =   "Color &dialog..."
      Height          =   375
      Left            =   45
      TabIndex        =   23
      Top             =   4005
      Width           =   2175
   End
   Begin VB.PictureBox pcMain 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   4635
      MouseIcon       =   "frmColorPicker.frx":0CCA
      MousePointer    =   99  'Custom
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   22
      ToolTipText     =   "Shift + click to make a gradient"
      Top             =   45
      Width           =   3900
   End
   Begin VB.PictureBox pcVertical 
      Height          =   3900
      Left            =   8640
      MouseIcon       =   "frmColorPicker.frx":0FD4
      MousePointer    =   99  'Custom
      Picture         =   "frmColorPicker.frx":12DE
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   21
      ToolTipText     =   "Shift + click to make a gradient"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton btPick 
      Caption         =   "&Pick from screen"
      Height          =   375
      Left            =   45
      TabIndex        =   9
      Top             =   3600
      Width           =   2175
   End
   Begin VB.PictureBox lbColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   645
      Left            =   585
      MouseIcon       =   "frmColorPicker.frx":4F22
      MousePointer    =   99  'Custom
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   37
      Top             =   45
      Width           =   3900
      Begin VB.PictureBox pcSmall 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   3
         Left            =   3645
         ScaleHeight     =   90
         ScaleWidth      =   90
         TabIndex        =   42
         ToolTipText     =   "Black"
         Top             =   405
         Width           =   120
      End
      Begin VB.PictureBox pcSmall 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   2
         Left            =   3645
         ScaleHeight     =   90
         ScaleWidth      =   90
         TabIndex        =   41
         ToolTipText     =   "White"
         Top             =   0
         Width           =   120
      End
      Begin VB.PictureBox pcSmall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   1
         Left            =   0
         ScaleHeight     =   90
         ScaleWidth      =   90
         TabIndex        =   40
         ToolTipText     =   "Invert color"
         Top             =   360
         Width           =   120
      End
      Begin VB.PictureBox pcSmall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   0
         Left            =   0
         ScaleHeight     =   90
         ScaleWidth      =   90
         TabIndex        =   39
         ToolTipText     =   "Current color"
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   10
      Left            =   4185
      MouseIcon       =   "frmColorPicker.frx":522C
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   9
      Left            =   3840
      MouseIcon       =   "frmColorPicker.frx":5536
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   8
      Left            =   3480
      MouseIcon       =   "frmColorPicker.frx":5840
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   7
      Left            =   3120
      MouseIcon       =   "frmColorPicker.frx":5B4A
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   6
      Left            =   2760
      MouseIcon       =   "frmColorPicker.frx":5E54
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   2400
      MouseIcon       =   "frmColorPicker.frx":615E
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   2040
      MouseIcon       =   "frmColorPicker.frx":6468
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   1680
      MouseIcon       =   "frmColorPicker.frx":6772
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   1320
      MouseIcon       =   "frmColorPicker.frx":6A7C
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   960
      MouseIcon       =   "frmColorPicker.frx":6D86
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   810
      Width           =   300
   End
   Begin VB.Label lbSlot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   600
      MouseIcon       =   "frmColorPicker.frx":7090
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Right-click to remember, left-click to retrieve"
      Top             =   810
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Slots:"
      Height          =   285
      Left            =   60
      TabIndex        =   10
      Top             =   825
      Width           =   465
   End
   Begin VB.Label lbGetColor 
      Caption         =   "Color:"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   510
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'*               BF Color Picker from BugFull Software                 *
'*           written by Chavdar Jordanov (chavo@beer.com)              *
'*     You may freely use and modify this code as long as you keep     *
'*                       this title intact.                            *
'*              Hope its gonna be useful for you!                      *
'***********************************************************************

'----- Note from the author: I deliberately did not use any API calls or C++ routines.
'-     Some of this code may work more efficiently with API or a C++ dll, but I am trying
'-     to show what can be done in pure Visual Basic. Good luck!
'----------------------------------------------------------------------------------

Option Explicit

Dim Col As Long                     'The main color
Dim bMouseOverPalette As Boolean    'Mouse is over the gradient palette
Dim OldX As Long

Const BigForm = 9200
Const SmallForm = 4700

'----- copies the color value to the Clipboard --------
Private Sub btCopy_Click(Index As Integer)
    Dim S As String
    S = txCol(Index).Text
    Clipboard.Clear
    Clipboard.SetText S
End Sub

'------ shows the windows color dialog ---------------
Private Sub btDialog_Click()
    On Error GoTo 100
    CDL.CancelError = True
    CDL.flags = cdlCCRGBInit
    CDL.Color = lbColor.BackColor
    CDL.ShowColor
    Col = CDL.Color
    'ShowColors (Col)
10
    Exit Sub
100
    Resume 10
End Sub

Private Sub btExit_Click()
 On Error GoTo 220

   
   SaveSetting App.Title, "Settings", "backcolor", frmColorPicker.BackColor
frmColorPicker.Visible = False
Form1.Visible = True
Form1.WindowState = vbMaximized
Form1.BackColor = frmColorPicker.BackColor
Form1.cmdRefresh.BackColor = frmColorPicker.BackColor
Form1.cmdStop.BackColor = frmColorPicker.BackColor
Form1.cmdBack.BackColor = frmColorPicker.BackColor
Form1.cmdForward.BackColor = frmColorPicker.BackColor
Form1.cmdGo.BackColor = frmColorPicker.BackColor
Form1.killwindow.BackColor = frmColorPicker.BackColor
Form1.Combo1.BackColor = frmColorPicker.BackColor
Form1.sldlbl.BackColor = frmColorPicker.BackColor
Form1.lblStatus.BackColor = frmColorPicker.BackColor
Form1.Combo2.BackColor = frmColorPicker.BackColor
Form1.cmdhome.BackColor = frmColorPicker.BackColor
Form1.Combo3.BackColor = frmColorPicker.BackColor
Form1.srchtext.BackColor = frmColorPicker.BackColor

220:
   SaveSetting App.Title, "Settings", "backcolor", vbRed
   
End Sub

'----------- unfolds or folds the form to show or hide the palette window -----
Private Sub btPalette_Click()
    If Me.Width = BigForm Then
        Me.Width = SmallForm
        btPalette.Caption = "&Palette >>>"
    Else
        Me.Width = BigForm
        btPalette.Caption = "&Palette <<<"
        If opPal(0) Then ShowGradientPalette Col
    End If
End Sub

'----------- captures the screen to frmScreen ------------
Private Sub btPick_Click()
    PrepareScreen
    btPick.Enabled = False
End Sub

'----------- shows a color from the IE color table ---------
Private Sub cbWeb_Click()
    If cbWeb.ListIndex > 0 Then
        Col = cbWeb.ItemData(cbWeb.ListIndex)
        ShowColors Col
    End If
End Sub

Private Sub chHue_Click()
    ShowGradientPalette Col
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        MsgBox "BF Color Picker (Freeware)" + vbCrLf + "BugFull Software 2001" + vbCrLf + "Written by Chavdar Yordanov" + vbCrLf + "E-mail: chavo@beer.com", vbInformation, "About BF Color Picker"
    End If
End Sub

Private Sub Form_Load()
    
    Col = RGB(255, 255, 255)
    ArrangeSmall             'arranges the small color slots within the lbColor
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 'center the form
    GetSlots                 'retrieves saved color values from registry
    GetWebColors Me.cbWeb    'load color values and names into the combo box
    iColorDepth = clr24Bit   'sets the default color depth
    opPal(1).Value = True
    Me.Width = SmallForm
    Me.Show
    
    ShowColors Col
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lbCol_Click(Index As Integer)
    btCopy_Click (Index)
End Sub

'------------- the main sub where it all takes place --------

Sub ShowColors(ByVal iCol As Long) 'Calculates the R<G<B values and writes them to the text boxes
    Dim R, G, B, i
    Dim bTmp(1 To 3) As Byte
    On Error Resume Next
    If iCol < 0 Then Exit Sub
    iCol = CalcColorDepth(iCol)
    lbColor.BackColor = iCol
    txCol(0).Text = CStr(iCol)
    'Split the long value into separate bytes
    SplitIntoBytes iCol, 3, bTmp, False
    'Assign the byte values to R,G,B variables just for convenience
    B = bTmp(3)
    G = bTmp(2)
    R = bTmp(1)
    
    lbRGB(0).BackColor = RGB(R, 0, 0)
    lbRGB(1).BackColor = RGB(0, G, 0)
    lbRGB(2).BackColor = RGB(0, 0, B)

    scColor(0).Value = R
    scColor(1).Value = G
    scColor(2).Value = B
    
    txCol(1) = "#" + Format(Hex(R), "00") + Format(Hex(G), "00") + Format(Hex(B), "00")
    txCol(2) = Format(R) + "," + Format(G) + "," + Format(B)
        
    pcSmall(0).BackColor = Col
    pcSmall(1).BackColor = Invert(Col)
    
    For i = 1 To cbWeb.ListCount - 1
        If cbWeb.ItemData(i) = Col Then
            cbWeb.ListIndex = i
            If opPal(2).Value Then ShowIEPalette Col
            Exit Sub
        End If
    Next i
    If opPal(1).Value Then ShowSafeSwatches Col
    cbWeb.ListIndex = 0
End Sub

Private Sub lbColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Col = lbColor.Point(x, y)
    ShowColors Col
    If Me.Width = BigForm And opPal(0).Value Then ShowGradientPalette Col
End Sub

'--------- sets or retrieves a color from the color slots ----------
Private Sub lbSlot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If Shift = 1 Then
            ShowGradient Col, lbSlot(Index).BackColor
        Else
            Col = lbSlot(Index).BackColor
            ShowColors Col
        End If
        If opPal(0) Then ShowGradientPalette Col
    Else
        lbSlot(Index).BackColor = lbColor.BackColor
        SaveSlots
    End If
End Sub

Private Sub opModify_Click(Index As Integer)
    iColorDepth = Choose(Index + 1, clr24Bit, clr16Bit, clrWebSafe)
    If Me.Width = BigForm And opPal(0).Value Then ShowGradientPalette Col
End Sub

'---------- shows one of the 3 available palettes -------------
Private Sub opPal_Click(Index As Integer)
    Select Case Index
    Case 0 'Gradient palette
        ShowGradientPalette Col
    Case 1 'Swatches 216 col
        ShowSafeSwatches
    Case 2 'swatches IE 4+
        ShowIEPalette
    End Select
'    chHue.Enabled = Index = 0
End Sub

Private Sub pcMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lCol As Long
    bMouseOverPalette = True
    lCol = pcMain.Point(x, y)
    If Shift = 1 Then
        ShowGradient Col, lCol
    Else
        ShowColors lCol
    End If
End Sub

Private Sub pcMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseOverPalette = False
End Sub

Private Sub pcSmall_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = 1 Then
        ShowGradient Col, pcSmall(Index).BackColor
    Else
        Col = pcSmall(Index).BackColor
        ShowColors Col
    End If
End Sub

Private Sub scColor_Change(Index As Integer)
frmColorPicker.BackColor = RGB(scColor(0), scColor(1), scColor(2))
    Col = RGB(scColor(0).Value, scColor(1).Value, scColor(2).Value)
    ShowColors Col
End Sub

Private Sub txCol_GotFocus(Index As Integer)
    SelectAll Index
End Sub

'---------- validates the input from the text boxes -----------
Private Sub txCol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SetColors Index
        SelectAll Index
        Exit Sub
    End If
    
    Dim sAllowed As String
    If KeyAscii > 31 Then
        Select Case Index
            Case 0
                sAllowed = "0123456789"
            Case 1
                sAllowed = "#0123456789abcdefABCDEF"
            Case 2
                sAllowed = "0123456789,"
        End Select
        If InStr(sAllowed, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

'--------- converts typed values to color --------------
Sub SetColors(iType As Integer)
    Dim sCol As String, N As Integer, i
    sCol = Condense(txCol(iType).Text)
    Col = 0
    On Error Resume Next
    Select Case iType
        Case 0 'Long
            Col = Val(sCol)
        Case 1 'Hex
            Col = HexToLong(sCol)
        Case 2 'RGB
            Col = RgbToLong(sCol)
    End Select
    ShowColors Col
End Sub

Function Condense(S As String) As String  'Removes all spaces from a string
    Dim i, C, Z
    Z = ""
    For i = 1 To Len(S)
        C = Mid(S, i, 1)
        If C <> " " Then Z = Z + C
    Next i
    Condense = Z
End Function

Sub SelectAll(Index As Integer)
    With txCol(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'---------- retrieves the color values for the slots from registry ----------
Sub GetSlots()
    Dim i
    For i = 0 To lbSlot.Count - 1
        lbSlot(i).Caption = " "
        lbSlot(i).ToolTipText = "Right-click to remember, left-click to retrieve"
        lbSlot(i).BackColor = GetSetting("BFColorPicker", "Slots", "Color" + CStr(i), vbWhite)
    Next i
End Sub
'---------- saves the color values to the registry --------------
Sub SaveSlots()
    Dim i
    For i = 0 To lbSlot.Count - 1
        SaveSetting "BFColorPicker", "Slots", "Color" + CStr(i), CStr(lbSlot(i).BackColor)
    Next i
End Sub

'=========== SCREEN CAPTURE FUNCTIONS =============
Private Sub PrepareScreen()
    Screen.MousePointer = 11
    If frmScreen.Visible = True Then
        Unload frmScreen
        Exit Sub
    Else
        'prepare frmScreen and capture the screen into picture1.
        frmScreen.Move 0, 0, Screen.Width, Screen.Height
        frmScreen.Picture1.Move 0, 0, frmScreen.Width, frmScreen.Height
        frmScreen.MousePointer = 99
        Set frmScreen.MouseIcon = lbColor.MouseIcon
        Set frmScreen.Picture1.Picture = CaptureScreen()
        frmScreen.Visible = True
    End If
    Screen.MousePointer = 0
End Sub

'============= COLOR PALETTE FUNCTIONS ================
Private Sub pcMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lCol As Long
    If bMouseOverPalette Then
        lCol = pcMain.Point(x, y)
        ShowColors lCol
    End If
End Sub
'----------- picks a color from the vertical palette on the right --------
Private Sub pcVertical_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lMaxColor As Long
    lMaxColor = pcVertical.Point(x, y)
    If Shift = 0 Then ShowColors lMaxColor Else ShowGradient Col, lMaxColor
    If opPal(0) Then ShowGradientPalette lMaxColor
    DoEvents
End Sub

'-------- shows the gradient palette (faster than the sub from the previous version which used PSET) -------
'-        takes about 0.5 seconds on my Athlon 600 machine
'-        Creates a bitmap file on the disk and then loads it into the picture box
'------------------------------------------------------------------

Public Sub ShowGradientPalette(ByVal lMaxColor As Long)
    Dim i As Long, j As Long              'Counters
    Dim R As Long, G As Long, B As Long        'Color values as bytes
    Dim k As Single                             'needed for the calculations
    Dim KF As Single                            'needed for the calculations
    Dim cPos As Long                            'Current position within the bitmap array
    Dim sFileName As String                     'Temporary name for the bitmap file
    Dim bBitmap(1 To 256 ^ 2 * 3 + 54) As Byte  'The array containing all the bitmap information to be saved to disk
    Dim bColorBytes(1 To 3) As Byte        'Holder for the RGB values
    Dim NewCol As Long
    Dim T
    Const bmpOffset = 54            'the header size for the bitmap disk file. Must skip it when loading color values into the bitmap array

    If lMaxColor < 0 Then Exit Sub  'Happens when user clicks on the picturebox borders
    Screen.MousePointer = 11
    SplitIntoBytes lMaxColor, 3, bColorBytes(), False
    R = bColorBytes(1)
    G = bColorBytes(2)
    B = bColorBytes(3)
    cPos = bmpOffset                'start writing color values after the file header
    T = Timer
    
    For i = 0 To 255
        KF = (i / 65025)
        For j = 255 To 0 Step -1
            k = (255 - j) * KF
            bColorBytes(1) = GetColorByte(k * B + j) 'CalcByte(B, i, j)       '(k * B + j)
            bColorBytes(2) = GetColorByte(k * G + j)
            bColorBytes(3) = GetColorByte(k * R + j)
            MergeBytes bBitmap(), bColorBytes(), cPos      'write the 3 byte color value to the bitmap array
        Next j
    Next i
    sFileName = "c:\cppal.bmp"                        'Assigns a temporary file for the bitmap palette
    Create24bitBitmap 256, 256, bBitmap(), sFileName  'creates a bitmap containg the palette on the harddisk
    Set pcMain.Picture = LoadPicture(sFileName)    'and loads it into pcMain
    Kill sFileName                                    'Delete the temporary file
    Screen.MousePointer = 0
    Debug.Print Int((Timer - T) * 1000)
End Sub



'----------- Shows a gradient between lMinCol and lMaxCol in lbColor ----------
Sub ShowGradient(ByVal lMinCol, ByVal lMaxCol)
    Dim i As Long, H, W
    Dim R1 As Long, r2 As Long, G1 As Long
    Dim g2 As Long, B1 As Long, b2 As Long
    Dim bBytes() As Byte
    Dim NewR As Long, NewB As Long, NewG As Long
    Dim NewCol As Long
    Dim Perc As Byte
    If lMinCol < 0 Or lMaxCol < 0 Then Exit Sub
    Screen.MousePointer = 11
    SplitIntoBytes lMaxCol, 3, bBytes()
    B1 = bBytes(3)
    G1 = bBytes(2)
    R1 = bBytes(1)
    SplitIntoBytes lMinCol, 3, bBytes()
    b2 = bBytes(3)
    g2 = bBytes(2)
    r2 = bBytes(1)
    
    lbColor.Cls
    lbColor.DrawMode = 13
    ShowColors lMaxCol
    H = lbColor.ScaleHeight
    W = lbColor.ScaleWidth
    lbColor.DrawMode = 13
    For i = 0 To 255
        NewR = i / 255 * R1 + (255 - i) / 255 * r2
        NewG = i / 255 * G1 + (255 - i) / 255 * g2
        NewB = i / 255 * B1 + (255 - i) / 255 * b2
        NewCol = CalcColorDepth(RGB(NewR, NewG, NewB))
        lbColor.Line (i, 0)-(i, H), NewCol
    Next i
    lbColor.DrawMode = 6
    lbColor.Line (0, H * 2 / 3)-(W, H * 2 / 3)
    Perc = 0
    For i = 0 To W Step W / 10
        lbColor.Line (i - 1, H * 2 / 3 - 5)-(i - 1, H * 2 / 3)
        lbColor.CurrentX = i - 6
        lbColor.CurrentY = H * 2 / 3 + 1
        lbColor.FontSize = 7
        lbColor.FontName = "Arial"
        If Perc > 0 Then lbColor.Print Perc
        Perc = Perc + 1
    Next i
    lbColor.DrawMode = 13
    Screen.MousePointer = 0
End Sub

'-------- arranges the small color slots in the right-top corner of the lbColor -----
Sub ArrangeSmall()
    Dim i
    For i = 1 To 4
        pcSmall(4 - i).Move lbColor.ScaleWidth - pcSmall(4 - i).Width * i, -1
    Next i
End Sub

'----------- shows the Internet Explorer color palette -------------
Private Sub ShowIEPalette(Optional ByVal ShowCol = -1)
    Dim HH, WW, i, j
    Dim Cnt As Integer
    Dim iCol As Long
    Dim bCol(140)
    For i = 1 To cbWeb.ListCount - 1
        bCol(i) = cbWeb.ItemData(i)
    Next i
    'SortArray bCol()
    With pcMain
        .Cls
        WW = .ScaleWidth / 12
        HH = .ScaleHeight / 12
        Cnt = 0
        For i = 0 To 11
            For j = 0 To 11
                Cnt = Cnt + 1
                If Cnt > 140 Then Exit For
                pcMain.Line (j * WW, i * HH)-((j + 1) * WW, (i + 1) * HH), bCol(Cnt), BF
                If ShowCol = bCol(Cnt) Then iCol = vbWhite Else iCol = vbBlack
                pcMain.Line (j * WW, i * HH)-((j + 1) * WW, (i + 1) * HH), iCol, B
            Next j
        Next i
    End With
End Sub

'---------------- shows 216 color palette ----------------
Private Sub ShowSafeSwatches(Optional ByVal ShowCol = -1)
    Dim HH, WW, i, j
    Dim Cnt As Integer
    Dim iCol As Long
    With pcMain
        .Cls
        WW = .ScaleWidth / 16
        HH = .ScaleHeight / 14
        Cnt = 0
        For i = 0 To 15
            For j = 0 To 13
                Cnt = Cnt + 1
                If Cnt > 224 Then Exit Sub
                pcMain.Line (i * WW, j * HH)-((i + 1) * WW, (j + 1) * HH), SafeCol(Cnt), BF
                If ShowCol = SafeCol(Cnt) Then iCol = vbWhite Else iCol = vbBlack
                pcMain.Line (i * WW, j * HH)-((i + 1) * WW, (j + 1) * HH), iCol, B
            Next j
        Next i
    End With
End Sub

