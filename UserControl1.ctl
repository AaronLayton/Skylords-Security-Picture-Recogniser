VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim imageFilter As String
Public Sub download(fileUrl As String)
    If Ambient.UserMode Then
        On Error Resume Next
        On Error GoTo errH
        AsyncRead fileUrl, vbAsyncTypePicture
    End If
    Exit Sub
errH:
    
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error Resume Next
    Image1.Picture = AsyncProp.Value
    imageFilter = Right(AsyncProp.Status, 4)
    Image1.Refresh
    UserControl.Height = Image1.Height
    UserControl.Width = Image1.Width + 50
    
    Form1.Picture1.Picture = Image1.Picture
End Sub


