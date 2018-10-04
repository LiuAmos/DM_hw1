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
Dim counter As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim class() As String
Dim gene() As String
Dim tvalue As String
Dim n1array(62) As Integer
Dim n2array(62) As Integer
Dim output, a
Set output = CreateObject("Scripting.Dictionary")

counter = -1
n1 = 0
n2 = 0
tvalue = "100"
Open App.Path & "\" + file For Input As #1

Do While Not EOF(1)
    Line Input #1, tmpline
    
    'class
    If counter = 0 Then
        class = Split(tmpline, ",")
        For i = 1 To 62
            'List1.AddItem class(i)
            If class(i) = "1" Then
                n1array(i) = i
                n1 = n1 + 1
            Else
                n2array(i) = i
            End If
        Next i
        n2 = 62 - n1
    
    Else
        If counter <> -1 Then
            gene = Split(tmpline, ",")
            output.Add gene(0), tvalue
        End If
    End If
    'List1.AddItem counter
    counter = counter + 1
Loop
Close #1

'List1.AddItem n1
'List1.AddItem n2
a = output.Keys
For i = 0 To output.Count - 1
    List1.AddItem a(i)
Next

End Sub

Private Sub Text1_Change()
file = Text1.Text
End Sub
