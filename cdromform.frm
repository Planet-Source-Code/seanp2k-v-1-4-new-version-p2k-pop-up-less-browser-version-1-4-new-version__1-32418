VERSION 5.00
Begin VB.Form cdromform 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pick your cd-rom drive......"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label label1 
      Caption         =   "Pick your cd-rom drive......."
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "cdromform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
Dim retval
Dim yourdrive As String
If retval <> 0 Then MsgBox "Not a CD drive!", vbOKOnly, "U R DUM"
cdromform.Visible = False
form1.WindowState = vbMaximized
retval = openCD(Mid(cdromform.Drive1.Drive, 1, 1))
If retval <> 0 Then MsgBox "Not a CD drive!", vbOKOnly, "U R DUM"
End Sub
