VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Skylords Security Pass"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3495
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   975
      Left            =   720
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Download New"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2415
      Left            =   4920
      TabIndex        =   3
      Top             =   2640
      Width           =   3375
      ExtentX         =   5953
      ExtentY         =   4260
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
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
      Location        =   "http:///"
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   720
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Split it"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3840
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   120
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Text2.Text = ""
    Text1.Text = ""
    List1.Clear
    col3 Picture1, 220
    Call Strip_Vert(Picture1)
End Sub

Private Sub Command2_Click()
    WebBrowser1.Refresh
End Sub

Private Sub Command3_Click()
    
    For x = 1 To 4
        If Mid(Text1.Text, x, 1) = "*" Then
            MsgBox "New Letter is " & Mid(Text2.Text, x, 1) & "," & List1.List(x - 1)
            
            Open App.Path & "\System.ipf" For Append As #2
                Write #2, vbNewLine & List1.List(x) & "," & Mid(Text2.Text, x, 1)
            Close #2
        End If
    Next x

End Sub

Private Sub Form_Load()
    WebBrowser1.Navigate2 "http://www.skylords.com/login"
    Picture1.Picture = LoadPicture(App.Path + "\Security\ch.jpg")
End Sub

Private Sub Timer1_Timer()
    col3 Picture1, 210
    Timer1.Enabled = False
    
End Sub


Private Sub WebBrowser1_DownloadComplete()
    Set Doc = WebBrowser1.Document
    
    Counter = 0
    For Each e In Doc.All.tags("IMG")
        Counter = Counter + 1
        
        If Counter = 9 Then
            UserControl11.download (e.src)
        End If
    Next
    
End Sub

