VERSION 5.00
Begin VB.Form warezfrm 
   Caption         =   "DO NOT CLOSE THIS WINDOW!!!!! CLICK DONE!! YOU WILL MESS IT UP!!!"
   ClientHeight    =   1395
   ClientLeft      =   -360
   ClientTop       =   855
   ClientWidth     =   9180
   Icon            =   "warezfrmp2k.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   9180
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6720
      Top             =   480
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto (might not work)"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   8280
      TabIndex        =   13
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate and download!"
      Height          =   315
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DONE"
      Height          =   195
      Left            =   5280
      TabIndex        =   9
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox s2 
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox s1 
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   7080
      TabIndex        =   6
      Text            =   "adds \/ to here"
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Text            =   "/files/"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "32"
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Text            =   "giants.C"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "group/ciw"
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "groups.yahoo.com/"
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Http://"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox bum2 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "This box will be hidden, so dont bother"
      Top             =   810
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox bum1 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "This box will be hidden, so dont bother"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "warezfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Check1_Click()
Select Case Check1
Case Checked
Timer1.Enabled = True
Case Unchecked
Timer1.Enabled = False
End Select
End Sub

Private Sub Command1_Click()
Form1.WindowState = vbMaximized
warezfrm.Hide
End Sub

Private Sub Command2_Click()
On Error Resume Next

Text5.Text = Text5.Text + Val(s2.Text)
Text7.Text = Text7.Text + Val(s2.Text)
Text7.Text = Text1.Text + Val(Text2.Text)
If Val(Text7.Text) < 9 Then GoTo fur
If Val(Text7.Text) > 9 Then Exit Sub
fur:
Text7.Text = "0" & Text7.Text

 Form1.WebBrowser1.Navigate (Text1.Text & Text2.Text & Text3.Text & Text5.Text & Text4.Text & Text6.Text & Text7.Text & Text8.Text)

End If

End Sub





Private Sub Command3_Click()
Help.Show
End Sub

Private Sub GO_Click()
urls.Text = Text1.Text & Text2.Text & Text3.Text & Val(Text5) + Val(s1) & Text4.Text & Text6.Text & Val(Text7) + Val(s2)
End Sub

Private Sub Command6_Click()
If DL1.CPause = True Or DL1.InDL = True Then Exit Sub
DL1.Url = "http://download.microsoft.com/download/vstudio60ent/SP5/Wideband-VB/WIN98Me/EN-US/vs6sp5vb.exe"
DL1.GetFileInformation
If DL1.FileSize <= 0 Then
Exit Sub
Else
lblconnected = "Connection Present - " & DL1.Connected
ProgressBar1.Max = DL1.FileSize
lblTotalBytes = "Total Bytes - " & DL1.FileSize
lblTS = "Total Size - " & DL1.FileSize
lblexists = "File Exists - " & DL1.FileExists
lblResume = "Resume Supported - " & DL1.AResume
DL1.DownLoad
End If
End Sub

Private Sub Text7_Change()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Form1.WebBrowser1.Navigate (Text1.Text & Text2.Text & Text3.Text & Text5.Text + Val(s1.Text) & Text4.Text & Text6.Text & (Text7.Text) + Val(s2.Text) & Text8.Text)
warezfrm.bum1.Text = Text5.Text + Val(s1)
warezfrm.bum2.Text = Text7.Text + Val(s2)
warezfrm.Text5.Text = bum1.Text
warezfrm.Text7.Text = bum2.Text
End Sub


