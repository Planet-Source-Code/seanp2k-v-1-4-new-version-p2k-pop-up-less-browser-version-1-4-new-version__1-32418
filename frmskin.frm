VERSION 5.00
Begin VB.Form frmskin 
   Caption         =   "Skin Making"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "frmskin.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "DONE"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Width           =   8295
   End
   Begin VB.TextBox Text1 
      Height          =   6015
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmskin.frx":0CCA
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmskin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmskin.Hide
Form1.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
Text1.Locked = True
End Sub
