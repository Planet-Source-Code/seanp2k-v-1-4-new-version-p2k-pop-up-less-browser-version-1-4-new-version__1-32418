VERSION 5.00
Begin VB.Form frmsplash 
   BorderStyle     =   0  'None
   Caption         =   "_+P2K+_ Browser"
   ClientHeight    =   7230
   ClientLeft      =   3375
   ClientTop       =   2835
   ClientWidth     =   9630
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   7260
      Left            =   0
      Picture         =   "frmsplash.frx":0CCA
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   0
      Width           =   9660
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   2880
         Top             =   4800
      End
      Begin VB.Timer t7 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2760
         Top             =   3720
      End
      Begin VB.Timer t6 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2280
         Top             =   3720
      End
      Begin VB.Timer t5 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1800
         Top             =   3720
      End
      Begin VB.Timer t4 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1320
         Top             =   3720
      End
      Begin VB.Timer t3 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   840
         Top             =   3720
      End
      Begin VB.Timer t2 
         Enabled         =   0   'False
         Interval        =   750
         Left            =   360
         Top             =   3720
      End
      Begin VB.Timer t1 
         Interval        =   1
         Left            =   0
         Top             =   3720
      End
      Begin VB.Timer Timer1 
         Interval        =   4750
         Left            =   2880
         Top             =   4200
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BorderColor     =   &H00808080&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   105
         Left            =   1920
         Top             =   60
         Width           =   15
      End
      Begin VB.Shape Shape1 
         Height          =   135
         Left            =   1905
         Top             =   50
         Width           =   7600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Starting _+P2K+_ Browser V 1.4.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   2
         Top             =   270
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1900
      End
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_Click()
MsgBox "Not done loading yet!!!", vbOKOnly, "NOT DONE LOADING YET!"
'IF YOU ARE A PROGRAMER, THE PASSWORD IS     prgm    SO YOU CAN SET THE TIMER DELAY FOR THE SPLASH SCREEN
End Sub

Private Sub t1_Timer()
Label2.Caption = "Loading Code"
t1.Enabled = False
t2.Enabled = True
End Sub

Private Sub t2_Timer()
Label2.Caption = "Loading Pictures"
t2.Enabled = False
t3.Enabled = True
End Sub

Private Sub t3_Timer()
Label2.Caption = "Loading Forms"
t3.Enabled = False
t4.Enabled = True
End Sub

Private Sub t4_Timer()
Label2.Caption = "Loading Buttons"
t4.Enabled = False
t5.Enabled = True
End Sub

Private Sub t5_Timer()
Label2.Caption = "Loading Modules"
t5.Enabled = False
t6.Enabled = True
End Sub

Private Sub t6_Timer()
Label2.Caption = "Loading Skins"

t6.Enabled = False
t7.Enabled = True
End Sub

Private Sub t7_Timer()
Label2.Width = 3495
Label2.Caption = "Starting _+P2K+_ Browser V " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Timer1_Timer()
Form1.Show
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
frmsplash.Height = Picture1.Height
frmsplash.Width = Picture1.Width
Timer1.Interval = GetSetting(App.Title, "settings", "timer1interval") ' you can change how long the splash screen stays up, it gets annoying if you are constantly running this in VB and have to wait a while.
Picture1.Picture = LoadPicture("Skin/SPLASH.JPG") 'loads the picture, but if its not there it has one built in
Timer1.Enabled = True

End Sub

Private Sub Timer2_Timer()
Label1.Caption = Text1.Text & " / 468 loaded"
Text1.Text = Text1.Text + 1
Shape2.Width = Shape2.Width + 16
If frmsplash.Text1.Text = 468 Then Timer2.Enabled = False
End Sub
