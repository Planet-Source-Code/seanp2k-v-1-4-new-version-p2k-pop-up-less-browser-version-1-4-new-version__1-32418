VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Searchfrm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Search........"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "PSCode"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Yahoo"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Profusion"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Google"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Type your query here then select a search engine"
      Top             =   0
      Width           =   4695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10935
      Left            =   0
      TabIndex        =   7
      Top             =   960
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
   Begin VB.Label Label1 
      Caption         =   "Search on........"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Searchfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub Command4_Click()
Form1.Combo3 = "http://www.profusion.com/searchresults.asp?queryterm=" + Text1.Text + "&AGT=Web&Category=1%2C1&CATID=1&option=all&RPP1=-1&rpe=30&totalverify=0&auto=all&Engine=349&E=349&Engine=1166&E=1166&Engine=1144&E=1144&Engine=1146&E=1146&Engine=1175&E=1175&Engine=1141&E=1141&Engine=1129&E=1129&Engine=1143&E=1143&Engine=354&E=354&Engine=363&E=363&Engine=1176&E=1176&Engine=1139&E=1139&Category=245%2C245%2C20%2Cxchg&SHW245=0&Category=6%2C6%2C20%2Cuser&SHW6=0&Category=172%2C172%2C20%2Cuser&SHW172=0&Category=96%2C96%2C20%2Cuser&SHW96=0"

End Sub

Private Sub Command6_Click()
 WebBrowser1.Navigate ("http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&B1=Quick+Search&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=" + Text1.Text + "&optSort=Alphabetical")
End Sub
