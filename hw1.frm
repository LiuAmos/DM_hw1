VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   13755
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "讀檔"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   5640
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   13455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file As String
Private Sub Command1_Click()
Open App.Path & "\" + file For Input As #1
Do While Not EOF(1)
          Line Input #1, tmpline
          List1.AddItem tmpline
Loop
Close #1
End Sub

Private Sub Text1_Change()
file = Text1.Text
End Sub
