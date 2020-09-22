VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Loading Homepage..... Please Wait....."
   ClientHeight    =   10095
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10890
   ForeColor       =   &H8000000E&
   Icon            =   "seanp2kbrowzer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox srchtext 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   5790
      TabIndex        =   17
      Text            =   "What to search for?"
      Top             =   240
      Width           =   2655
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "seanp2kbrowzer.frx":0CCA
      Left            =   8400
      List            =   "seanp2kbrowzer.frx":0CDD
      TabIndex        =   16
      Text            =   "What search engine?"
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdhome 
      BackColor       =   &H000000FF&
      Caption         =   "Home"
      Height          =   250
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   5775
   End
   Begin VB.CommandButton killwindow 
      BackColor       =   &H000000FF&
      Caption         =   "Clear history and close window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4820
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   220
      Width           =   975
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H000000FF&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2890
      Picture         =   "seanp2kbrowzer.frx":0D38
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   220
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Type in the URL's here"
      Top             =   890
      Width           =   5175
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H000000FF&
      Caption         =   "Go!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3840
      Picture         =   "seanp2kbrowzer.frx":12FA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H000000FF&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1920
      Picture         =   "seanp2kbrowzer.frx":1AA8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   220
      Width           =   975
   End
   Begin VB.CommandButton cmdForward 
      BackColor       =   &H000000FF&
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   960
      Picture         =   "seanp2kbrowzer.frx":21EA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   220
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      Picture         =   "seanp2kbrowzer.frx":2627
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   220
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
      ExtentX         =   18018
      ExtentY         =   19288
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Text            =   "Input your search phrase in here"
      ToolTipText     =   "Input the search phrase here"
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   300
      Left            =   840
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Text            =   "Bookmarks"
      ToolTipText     =   "Browse you bookmarks here"
      Top             =   890
      Visible         =   0   'False
      Width           =   4815
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   555
      Left            =   5880
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   979
      _Version        =   327682
      Max             =   255
   End
   Begin VB.Label sldlbl 
      BackColor       =   &H000000FF&
      Caption         =   $"seanp2kbrowzer.frx":2ADA
      Height          =   735
      Left            =   5880
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   0
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu NewWindow 
         Caption         =   "New Window"
         Shortcut        =   ^N
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Properties 
         Caption         =   "Properties"
      End
      Begin VB.Menu PageSetp 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu Print 
         Caption         =   "Print"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu SelectAll 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu hmpg 
      Caption         =   "Homepage"
      Begin VB.Menu mk_cur_home 
         Caption         =   "Make current page my homepage"
      End
      Begin VB.Menu set_home 
         Caption         =   "Set homepage"
      End
   End
   Begin VB.Menu Bookmarks 
      Caption         =   "Bookmarks"
      Begin VB.Menu show_book 
         Caption         =   "Show bookmarks"
      End
      Begin VB.Menu hide_book 
         Caption         =   "Hide bookmarks"
      End
      Begin VB.Menu ClearBkmarks 
         Caption         =   "Clear Bookmarks"
      End
      Begin VB.Menu Add 
         Caption         =   "Add a Bookmark"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Popups 
         Caption         =   "Disable Popups"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
      Begin VB.Menu Clear 
         Caption         =   "Clear History"
      End
   End
   Begin VB.Menu search 
      Caption         =   "Search"
      Begin VB.Menu hide_srch 
         Caption         =   "Hide searching stuff"
      End
      Begin VB.Menu show_srch 
         Caption         =   "Show searching stuff"
      End
   End
   Begin VB.Menu Visibility_menu 
      Caption         =   "Transparency"
      Begin VB.Menu neunf 
         Caption         =   "(as in make this window so you can see through it, WIN XP OR 2K ONLY)"
      End
      Begin VB.Menu VIS_MENU 
         Caption         =   "Visibility Slider (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu vis10 
         Caption         =   "Visibility 10% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu vis20 
         Caption         =   "Visibility 20% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu vis30 
         Caption         =   "Visibility 30% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu vis40 
         Caption         =   "Visibility 40% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu vis50 
         Caption         =   "Visibility 50% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu vis60 
         Caption         =   "Visibility 60% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu vis70 
         Caption         =   "Visibility 70% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu vis80 
         Caption         =   "Visibility 80% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu vis90 
         Caption         =   "Visibility 90% (WINDOWS XP OR 2K ONLY)"
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu vis100 
         Caption         =   "Visibility 100%"
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu colors 
      Caption         =   "Colors"
      Begin VB.Menu col_selector 
         Caption         =   "Color Selector"
      End
   End
   Begin VB.Menu fun_and_misc 
      Caption         =   "Fun and Misc"
      Begin VB.Menu how_long_splash_scrn 
         Caption         =   "How long the splash screen stays up for"
      End
      Begin VB.Menu warez_dl 
         Caption         =   "Warez downloader"
      End
      Begin VB.Menu nice_acsii 
         Caption         =   "WOW! A really REALLY nice image to ASCII program, writen by Arvinder Sehmi"
      End
      Begin VB.Menu start_max 
         Caption         =   "Start up maximized"
      End
      Begin VB.Menu start_min 
         Caption         =   "Start up minimized"
      End
      Begin VB.Menu start_norm 
         Caption         =   "Start up normal"
      End
      Begin VB.Menu Custom_title 
         Caption         =   "Custom Title"
      End
      Begin VB.Menu click_this_cupholder 
         Caption         =   "Click this for a FREE cup holder (joke)"
      End
      Begin VB.Menu who_made 
         Caption         =   "Who made this?"
      End
   End
   Begin VB.Menu makeaskinmenu 
      Caption         =   "Help and Skin making"
      Begin VB.Menu skin_make_how 
         Caption         =   "How do I make a skin for this?"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AllowPopups As Boolean
Public NumberOfTimesClicked As Integer
Public Index As Integer
Public State As Integer
Dim mbDontNavigateNow As Boolean
    

Private Sub Add_Click()
    AddBookmarks (App.Path & "\bookmarks.txt")
End Sub


Private Sub Clear_Click()
    On Error Resume Next
    Combo1.Clear
    Combo1.Text = WebBrowser1.LocationURL
    Dim i As Integer
    Dim a As String
    Open App.Path & "\history.txt" For Output As #1
    For i = 0 To Combo1.ListCount - 1
    Write #1, ""
    Next i
    Close #1
End Sub
Private Sub ClearBkmarks_Click()
    On Error Resume Next
    Combo2.Clear
    Dim i As Integer
    Dim a As String
    Open App.Path & "\bookmarks.txt" For Output As #1
    For i = 0 To Combo2.ListCount - 1
    Write #1, ""
    Next i
    Close #1
    Combo2.Text = "Bookmarks"
End Sub

Private Sub click_this_cupholder_Click()
cdromform.Visible = True
Form1.WindowState = vbNormal
End Sub


Private Sub cmbSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdSearch.Default = True
End Sub

Private Sub cmdhome_Click()
On Error Resume Next
WebBrowser1.Navigate (GetSetting(App.Title, "settings", "homepage", "www.google.com"))
End Sub

Private Sub cmdStop_Click()
    WebBrowser1.Stop
    ProgressBar1.Visible = False
    WebBrowser1.Height = Form1.ScaleHeight - 1520
    lblStatus.Caption = "Loading stopped"
    Combo1.SetFocus
End Sub

Private Sub cmdGo_Click()
    WebBrowser1.Navigate (Combo1.Text)
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    WebBrowser1.Refresh
End Sub

Private Sub col_selector_Click()
Form1.WindowState = vbNormal
frmColorPicker.Visible = True
End Sub

Private Sub Combo2_Change()
    WebBrowser1.Navigate Combo2.SelText
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    WebBrowser1.Navigate ("http://search.dogpile.com/texis/search?q=" & cmbSearch.Text & "&geo=no&refer=dp-search&fs=web")
    cmbSearch.AddItem (cmbSearch.Text)
    cmbSearch.SetFocus
End Sub

Private Sub config_cdrom_Click()

End Sub

Private Sub Command1_Click()
End Sub

Private Sub Combo3_Click()
Select Case Combo3.Text
Case "Google"
WebBrowser1.Navigate ("http://www.google.com/search?as_q=" + srchtext.Text + "&num=100&btnG=Google+Search&as_epq=&as_oq=&as_eq=&lr=&as_ft=i&as_filetype=&as_qdr=all&as_occt=any&as_dt=i&as_sitesearch=&safe=off")
Case "Profusion"
WebBrowser1.Navigate ("http://www.profusion.com/searchresults.asp?queryterm=" + srchtext.Text + "&AGT=Web&Category=1%2C1&CATID=1&option=all&RPP1=-1&rpe=30&totalverify=0&auto=all&Engine=349&E=349&Engine=1166&E=1166&Engine=1144&E=1144&Engine=1146&E=1146&Engine=1175&E=1175&Engine=1141&E=1141&Engine=1129&E=1129&Engine=1143&E=1143&Engine=354&E=354&Engine=363&E=363&Engine=1176&E=1176&Engine=1139&E=1139&Category=245%2C245%2C20%2Cxchg&SHW245=0&Category=6%2C6%2C20%2Cuser&SHW6=0&Category=172%2C172%2C20%2Cuser&SHW172=0&Category=96%2C96%2C20%2Cuser&SHW96=0")
Case "Yahoo"
WebBrowser1.Navigate ("http://search.yahoo.com/search?p=" + srchtext.Text + "&w=dir&fr=op&o=a&h=&g=0&n=100")
Case "Dogpile"
WebBrowser1.Navigate ("http://search.dogpile.com/texis/search?q=" + srchtext.Text + "&geo=no&fs=web")
Case "PSC"
WebBrowser1.Navigate ("http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&B1=Quick+Search&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=" + srchtext.Text + "&optSort=Alphabetical")
End Select
End Sub

Private Sub Copy_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Custom_title_Click()
titlefrm.Visible = True
End Sub

Private Sub Cut_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Exit_Click()
Unload Me
End
End Sub

Public Function coolClose(FormClose As Form, speed As Integer)
Me.WindowState = vbNormal
Do Until FormClose.Height <= 405
DoEvents
FormClose.Height = FormClose.Height - speed * 9
FormClose.Top = FormClose.Top + speed * 5
Loop
Do Until FormClose.Width <= 1680
DoEvents
FormClose.Width = FormClose.Width - speed * 9
FormClose.Left = FormClose.Left + speed * 5
Loop
Unload FormClose
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
'IF YOU ARE A PROGRAMER, THE PASSWORD IS     prgm    SO YOU CAN SET THE TIMER DELAY FOR THE SPLASH SCREEN

Private Sub Form_Load()
On Error Resume Next
Form1.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
cmdhome.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
cmdRefresh.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
cmdStop.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
cmdBack.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
cmdForward.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
cmdGo.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
killwindow.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
Combo1.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
sldlbl.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
lblStatus.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
Combo3.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
srchtext.BackColor = GetSetting(App.Title, "Settings", "backcolor", vbRed)
On Error Resume Next
cmdRefresh.Picture = LoadPicture("Skin\REFRESH.GIF")
cmdStop.Picture = LoadPicture("Skin\STOP.GIF")
cmdBack.Picture = LoadPicture("Skin\BACK.GIF ")
cmdForward.Picture = LoadPicture("Skin\FORWARD.GIF")
cmdGo.Picture = LoadPicture("Skin\GO.GIF")
Form1.Picture = LoadPicture("Skin\MAIN.GIF")

    
    
   
    
    State = 1
    WebBrowser1.Navigate (GetSetting(App.Title, "settings", "homepage", "www.google.com"))
    NumberOfTimesClicked = 0
    AllowPopups = False
    LoadBookmarks
    Dim a As String
    On Error Resume Next
    Open App.Path & "\history.txt" For Input As #1
    Do
        Input #1, a
        If a <> "" Then
        Combo1.AddItem a
    End If
    Loop Until EOF(1)
    Close #1
    Form1.WindowState = GetSetting(App.Title, "Settings", "state", vbMaximized)
errcolor:
   Resume Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Form1.WindowState <> 1 Then
        WebBrowser1.Width = Form1.ScaleWidth
        WebBrowser1.Height = Form1.ScaleHeight - 1525
        ProgressBar1.Width = Form1.ScaleWidth
        ProgressBar1.Top = Form1.ScaleHeight - 250
        lblStatus.Width = Form1.ScaleWidth - Combo1.Width
        cmbSearch.Width = Form1.Width - Combo1.Width - cmdSearch.Width - 300
        cmdSearch.Left = Form1.ScaleWidth - cmdSearch.Width
    End If
End Sub

Private Sub cmdBack_Click()
    On Error Resume Next
    WebBrowser1.GoBack
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub LoadBookmarks()
    Dim a As String
    On Error Resume Next
    Open App.Path & "\bookmarks.txt" For Input As #1
    Do
    Input #1, a
    If a <> "" Then
    Combo2.AddItem a
    End If
    Loop Until EOF(1)
    Close #1
End Sub

Private Sub Home_Click()
    WebBrowser1.GoHome
End Sub

Private Sub hide_book_Click()
Combo2.Visible = False
End Sub

Private Sub hide_srch_Click()
Combo3.Visible = False
srchtext.Visible = False
End Sub

Private Sub how_long_splash_scrn_Click()
frmspcfg.Show vbModal
End Sub

Private Sub killwindow_Click()
 On Error Resume Next
    Combo1.Clear
    Combo1.Text = WebBrowser1.LocationURL
    Dim i As Integer
    Dim a As String
    Open App.Path & "\history.txt" For Output As #1
    For i = 0 To Combo1.ListCount - 1
    Write #1, ""
    Next i
    Close #1
Unload Me
End
End Sub

Private Sub mk_cur_home_Click()
SaveSetting App.Title, "settings", "homepage", Form1.WebBrowser1.LocationURL
End Sub

Private Sub NewWindow_Click()
    On Error Resume Next
    Static lDocumentCount As Long
    Dim frmD As Form
    lDocumentCount = lDocumentCount + 1
    Set frmD = New Form1
    frmD.Show
    frmD.SetFocus
End Sub

Private Sub AddBookmarks(Filename As String)
    Combo2.AddItem WebBrowser1.LocationURL
    Dim i As Integer
    Dim a As String
    Dim Url As String
    Open App.Path & "\bookmarks.txt" For Output As #1
    For i = 0 To Combo1.ListCount + 1
    Write #1, Combo2.List(i)
    Next i
    Close #1
End Sub

Private Sub nice_acsii_Click()
ImageFrm.Visible = True
MainFrm.Visible = True
OptionsFrm.Visible = True
End Sub

Private Sub PageSetp_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Paste_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub Popups_Click()
    If Popups.Checked = False Then
        Popups.Checked = True
        AllowPopups = False
    ElseIf Popups.Checked = True Then
        Popups.Checked = False
        AllowPopups = True
    End If
End Sub

Private Sub Print_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub ProgressBar1_change()
End Sub

Private Sub Properties_Click()
    WebBrowser1.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Save_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub search_Click()
Combo3.Visible = False
End Sub

Private Sub SelectAll_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub set_home_Click()
homepage.Show vbModal
End Sub

Private Sub show_book_Click()
Combo2.Visible = True
End Sub

Private Sub show_srch_Click()
Combo3.Visible = True
srchtext.Visible = True
End Sub

Private Sub skin_make_how_Click()
Form1.WindowState = vbNormal
frmskin.Visible = True
End Sub

Private Sub Slider1_Change()
MakeTransparent Me.hWnd, Form1.Slider1.Value
End Sub


Private Sub start_max_Click()
 SaveSetting App.Title, "Settings", "state", vbMaximized
 Form1.WindowState = vbMaximized
End Sub

Private Sub start_min_Click()
 SaveSetting App.Title, "Settings", "state", vbMinimized
  Form1.WindowState = vbMinimized
End Sub

Private Sub start_norm_Click()
 SaveSetting App.Title, "Settings", "state", vbNormal
  Form1.WindowState = vbNormal
End Sub

Private Sub VIS_MENU_click()
  If VIS_MENU.Checked = False Then
        VIS_MENU.Checked = True
        Form1.Slider1.Visible = True
        Form1.sldlbl.Visible = True
        Form1.Combo2.Visible = False
        srchtext.Visible = False
        Combo3.Visible = False
    ElseIf VIS_MENU.Checked = True Then
        VIS_MENU.Checked = False
        Form1.Slider1.Visible = False
        Form1.sldlbl.Visible = False
        Form1.Combo2.Visible = True
       End If
End Sub
Private Sub vis_slid_off_Click()
Form1.Slider1.Visible = False
Form1.sldlbl.Visible = False
End Sub

Private Sub vis10_Click()
MakeTransparent Me.hWnd, 52.5
End Sub

Private Sub vis100_Click()
MakeTransparent Me.hWnd, 255
End Sub

Private Sub vis20_Click()
MakeTransparent Me.hWnd, 75
End Sub

Private Sub vis30_Click()
MakeTransparent Me.hWnd, 97.5
End Sub

Private Sub vis40_Click()
MakeTransparent Me.hWnd, 120
End Sub

Private Sub vis50_Click()
MakeTransparent Me.hWnd, 142.5
End Sub

Private Sub vis60_Click()
MakeTransparent Me.hWnd, 165
End Sub

Private Sub vis70_Click()
MakeTransparent Me.hWnd, 187.5
End Sub

Private Sub vis80_Click()
MakeTransparent Me.hWnd, 210
End Sub

Private Sub vis90_Click()
MakeTransparent Me.hWnd, 232.5
End Sub

Private Sub warez_dl_Click()
warezfrm.Show
Form1.WindowState = vbNormal
End Sub

Private Sub WebBrowser1_DownloadBegin()
    ProgressBar1.Visible = True
    WebBrowser1.Height = Form1.ScaleHeight - 1855
    ProgressBar1.Max = 1
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
    Combo1.Text = WebBrowser1.LocationURL
   Form1.Caption = GetSetting(App.Title, "Settings", "title") + WebBrowser1.LocationName
    On Error GoTo 123
  
    
    ProgressBar1.Visible = False
    WebBrowser1.Height = Form1.ScaleHeight - 1520
    Combo1.AddItem Combo1.Text
    Dim i As Integer
    Dim a As String
    Open App.Path & "\history.txt" For Output As #1
    For i = 0 To Combo1.ListCount - 1
    Write #1, Combo1.List(i)
    Next i
    Close #1
    cmdGo.Default = True
123
    SaveSetting App.Title, "Settings", "title", "_+Seanp2k Anti-Pop-up Browser 1.4 -"
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    If AllowPopups = True Then
        Cancel = False
        DoEvents
    ElseIf AllowPopups = False Then
        Cancel = True
    End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    If ProgressMax >= 0 And Progress > 0 And Progress <= ProgressMax Then
        ProgressBar1.Value = Progress / ProgressMax
    End If
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    lblStatus.Caption = Text
End Sub

Private Sub cmdForward_Click()
    On Error Resume Next
    WebBrowser1.GoForward
End Sub

Private Sub who_made_Click()
MsgBox "_+Seanp2k+_ made this, see www.extracredit.da.ru, the scroll down, for the latest version of this.", vbOKOnly, "Who made this?"
End Sub
