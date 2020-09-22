VERSION 5.00
Begin VB.Form frmspcfg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configure the splash screen"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Save and exit"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "3000"
      Top             =   1320
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   0
      TabIndex        =   1
      Text            =   "Password"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Enter password (found in source code of this form)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "How long (in milliseconds) should the splash screen stay up for?"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "frmspcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveSetting App.Title, "settings", "timer1interval", Text2.Text
frmspcfg.Hide
End Sub

Private Sub Command2_Click()
'if you can see this, the passowrd is prgm
On Error Resume Next
If Text1.Text = "prgm" Then GoTo 350 Else GoTo usuk
GoTo endlin
usuk:
MsgBox "If you are a programmer, go to www.extracredit.da.ru and look for the source for the p2k browser and you can download a copy of my source code for this, then look in the form load section of the code for this form, and a comment will be there with the password in it. If the source there contains no password, then go to pscode.com and search for p2k and download that source. If you are not a programmer, YOU CANNOT MODIFY THE AMOUNT OF TIME THE SPLASH SCREEN STAYS UP FOR!!!", vbOKOnly, "You got it wrong"
Label1.Visible = False
Command1.Visible = False
Text2.Visible = False
GoTo endlin

350:

Label1.Visible = True
Command1.Visible = True
Text2.Visible = True
MsgBox "Hello, fellow prgm'r", vbOKOnly, "HELLO"
GoTo endlin

endlin:
End Sub
