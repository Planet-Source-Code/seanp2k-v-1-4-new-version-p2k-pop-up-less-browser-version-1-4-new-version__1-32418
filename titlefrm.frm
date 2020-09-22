VERSION 5.00
Begin VB.Form titlefrm 
   Caption         =   "What do you want the title to be?"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "titlefrm.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   930
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Insert *-* and *space*"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Default"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Text            =   "_+Seanp2k Anti-Pop-up Browser 1.39 - "
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "NOTE: the page title will be displayed after what you type in."
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "titlefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveSetting App.Title, "Settings", "title", Text1.Text
Form1.Caption = titlefrm.Text1.Text
titlefrm.Visible = False
Form1.WindowState = vbMaximized
Form1.Caption = GetSetting(App.Title, "Settings", "title") + Form1.WebBrowser1.LocationName
End Sub

Private Sub Command2_Click()
Text1.Text = "_+Seanp2k Anti-Pop-up Browser 1.39 -"
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text + " - "
End Sub
