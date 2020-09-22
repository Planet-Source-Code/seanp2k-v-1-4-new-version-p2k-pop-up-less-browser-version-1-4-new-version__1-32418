VERSION 5.00
Begin VB.Form homepage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set your homepage"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   Icon            =   "homepage.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Use current"
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save and exit"
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "homepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveSetting App.Title, "settings", "homepage", Text1.Text
homepage.Hide
End Sub

Private Sub Command2_Click()
Text1.Text = Form1.WebBrowser1.LocationURL

End Sub

