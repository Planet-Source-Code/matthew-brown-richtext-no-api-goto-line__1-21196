VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goto Line Function Example"
   ClientHeight    =   3375
   ClientLeft      =   2670
   ClientTop       =   4215
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6975
   Begin VB.CommandButton Command1 
      Caption         =   "&Go!"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "1"
      Top             =   2760
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Goto line:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Enter some text here:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Yes! By demand... An example showing you
' how to use the SetCursorAtLine function
'
' I thought it was pretty easy to figure out
' but, anyway...
'
' Put some text in the RichText box,
' enter a line number in the Goto Line box
' and click 'Go!'
'
' You call the function 'SetCursorAtLine'
' passing it the line number and the RichText
' control respectively as arguments.
'
' See the bas for more info...
'
' -- Matthew Brown

Private Sub Command1_Click()
' Call the function
SetCursorAtLine Val(Text1), RichTextBox1
' So you can see the cursor!
RichTextBox1.SetFocus
End Sub


Private Sub Form_Load()
Me.Move (Screen.Width - Width) / 2, _
    (Screen.Height - Height) / 2
' Fill text box
For X = 1 To 10
    RichTextBox1.Text = RichTextBox1.Text & "Line " & X & vbCrLf
Next

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
' only allow numbers!
    Select Case Chr(KeyAscii)
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
    Case Else
        If KeyAscii = 13 Then
            Command1.Value = True
        ElseIf KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End Select
End Sub


