VERSION 5.00
Begin VB.Form warezfrm 
   Caption         =   "DO NOT CLOSE THIS WINDOW!!!!! CLICK DONE!! YOU WILL MESS IT UP!!!"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "HELP!!!!!!!"
      Height          =   315
      Left            =   4680
      TabIndex        =   17
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   8280
      TabIndex        =   14
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox numloop 
      Height          =   615
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   11
      Text            =   "warezfrm.frx":0000
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate and prepare for download!"
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
      TabIndex        =   13
      Text            =   "This box will be hidden, so dont bother"
      Top             =   810
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox bum1 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "warezfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Command1_Click()

warezfrm.Hide
End Sub

Private Sub Command2_Click()
On Error Resume Next
Form1.cboURL.Clear

Do Until Form1.cboURL.ListCount = Val(numloop.Text)
Form1.cboURL.AddItem (Text1.Text & Text2.Text & Text3.Text & Val(Text5) + Val(s1) & Text4.Text & Text6.Text & Val(Text7) + Val(s2))
warezfrm.bum1.Text = warezfrm.Text5.Text + Val(warezfrm.s1.Text)
warezfrm.bum2.Text = warezfrm.Text7.Text + Val(warezfrm.s2.Text)
warezfrm.Text5.Text = warezfrm.bum1.Text
warezfrm.Text7.Text = warezfrm.bum2.Text
Loop
End Sub





Private Sub Command3_Click()
Help.Show
End Sub

Private Sub GO_Click()
urls.Text = Text1.Text & Text2.Text & Text3.Text & Val(Text5) + Val(s1) & Text4.Text & Text6.Text & Val(Text7) + Val(s2)
End Sub

Private Sub Timer1_Timer()
Counter = Counter + 1
Text10.Text = Counter
End Sub

