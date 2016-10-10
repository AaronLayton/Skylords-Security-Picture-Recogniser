VERSION 5.00
Begin VB.Form addChar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter the character that this picture represents!"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   2655
   End
End
Attribute VB_Name = "addChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "Please enter the letter!"
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
End Sub
