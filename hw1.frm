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
Dim m, k, p, q As Integer
Dim tmp As Double


Private Sub Command1_Click()

'declare variable
Dim counter As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim class() As String
Dim gene() As String
Dim Dblgene(62) As Double
Dim Egene() As String
Dim ten As Double
Dim n1array(62) As Integer
Dim n2array(62) As Integer
Dim sum1 As Double
Dim sum2 As Double
Dim x1 As Double
Dim x2 As Double
Dim vsum1 As Double
Dim vsum2 As Double
Dim var1 As Double
Dim var2 As Double
Dim t As Double
Dim output, ke, it
Dim outputarray(2000) As Double
Set output = CreateObject("Scripting.Dictionary")

'assign value
counter = -1
n1 = 0
n2 = 0
ten = 10

Open App.Path & "\" + file For Input As #1

Do While Not EOF(1)
    Line Input #1, tmpline
    
    sum1 = 0
    sum2 = 0
    vsum1 = 0
    vsum2 = 0

    If counter = 0 Then
        class = Split(tmpline, ",")
        For i = 1 To 62
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
            For i = 0 To 62
                gene(i) = Trim(gene(i))
                If InStr(gene(i), "E") <> 0 Then
                    Egene = Split(gene(i), "E")
                    Dblgene(i) = CDbl(CDbl(Egene(0)) * (ten ^ CDbl(-Right(Egene(1), 1))))
                Else
                    Dblgene(i) = CDbl(gene(i))
                End If
            Next
            
            For i = 1 To 62
                If n1array(i) <> 0 Then
                    sum1 = sum1 + Dblgene(i)
                End If
                
                If n2array(i) <> 0 Then
                    sum2 = sum2 + Dblgene(i)
                End If
            Next
            
            x1 = (sum1 / CDbl(n1))
            x2 = (sum2 / CDbl(n2))
            
            For i = 1 To 62
                If n1array(i) <> 0 Then
                    vsum1 = vsum1 + (Dblgene(i) - x1) ^ 2
                End If
                
                If n2array(i) <> 0 Then
                    vsum2 = vsum2 + (Dblgene(i) - x2) ^ 2
                End If
            Next
            
            var1 = vsum1 / (n1 - 1)
            var2 = vsum2 / (n2 - 1)
            
            t = (x1 - x2) / ((var1 / n1) + (var2 / n2)) ^ 0.5
            
            
            If counter > 0 Then
                outputarray(counter) = t
            End If
            output.Add CDbl(gene(0)), t
        End If
    End If
    
    counter = counter + 1
    
Loop
Close #1


For k = 1 To 2000
    For m = k To 2000
        If outputarray(k) > outputarray(m) Then
            tmp = outputarray(k)
            outputarray(k) = outputarray(m)
            outputarray(m) = tmp
        End If
    Next m
Next k

For p = 1 To 2000
    For q = 0 To 1999
        If output.Item(q) = outputarray(p) Then
            List1.AddItem "Gene seq  " & CStr(q) & vbTab & "t =   " & CStr(outputarray(p))
            
        End If
    Next q
Next p



End Sub

Private Sub Text1_Change()
file = Text1.Text
End Sub
