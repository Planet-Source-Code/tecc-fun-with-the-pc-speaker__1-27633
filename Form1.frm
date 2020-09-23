VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fun with the PC Speaker"
   ClientHeight    =   5280
   ClientLeft      =   1560
   ClientTop       =   1170
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9015
   Begin VB.Frame Frame4 
      Caption         =   "Pictograph"
      Height          =   3855
      Left            =   5040
      TabIndex        =   22
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox Picture1 
         Height          =   3015
         Left            =   120
         ScaleHeight     =   2955
         ScaleWidth      =   3555
         TabIndex        =   24
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   23
         Text            =   "1"
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   3480
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Presets"
      Height          =   1095
      Left            =   5040
      TabIndex        =   20
      Top             =   4080
      Width           =   3615
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Effects"
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   4815
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   240
         TabIndex        =   17
         Text            =   "1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "100"
      Top             =   3240
      Width           =   3135
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   4335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "GA"
      Height          =   1215
      Left            =   4320
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "G"
      Height          =   1215
      Left            =   3960
      TabIndex        =   10
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "FG"
      Height          =   1215
      Left            =   3600
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "F"
      Height          =   1215
      Left            =   3240
      TabIndex        =   8
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "E"
      Height          =   1215
      Left            =   2880
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "DE"
      Height          =   1215
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "D"
      Height          =   1215
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CD"
      Height          =   1215
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "C"
      Height          =   1215
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "B"
      Height          =   1215
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AB"
      Height          =   1215
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Keyboard"
      Height          =   3495
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Ge 
         Caption         =   "Generate Silence"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Length MS:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Beep 55 * (List1.ListIndex + 1), Text1.Text

Select Case List2.ListIndex
Case 1
For e = 55 To (55 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = 55 To (55 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = 55 To (55 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = 55 To (55 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = 55 To (55 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = 55 To (55 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub cnote_Click(Index As Integer)


End Sub

Private Sub Command10_Click()
Dim val1 As Integer
val1 = 92
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command11_Click()
Dim val1 As Integer
val1 = 98
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command12_Click()
Dim val1 As Integer
val1 = 103
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command13_Click()
Select Case List2.ListIndex
Case 0
LongSweep
End Select
End Sub

Private Sub Command2_Click()

Dim val1 As Integer
val1 = 58
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command3_Click()
Dim val1 As Integer
val1 = 62
Beep val1 * (List1.ListIndex + 1), Text1.Text

Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command4_Click()
Dim val1 As Integer
val1 = 65
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command5_Click()
Dim val1 As Integer
val1 = 69
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command6_Click()
Dim val1 As Integer
val1 = 73
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command7_Click()
Dim val1 As Integer
val1 = 78
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command8_Click()
Dim val1 As Integer
val1 = 82
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Command9_Click()
Dim val1 As Integer
val1 = 87
Beep val1 * (List1.ListIndex + 1), Text1.Text
Select Case List2.ListIndex
Case 1
For e = val1 To (val1 - 10) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 2
For ep = val1 To (val1 + 10) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
Case 3
For ep = val1 To (val1 + Int(Rnd * 100)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(Rnd * 100)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case 4
For ep = val1 To (val1 + Int(10)) Step 1
Beep ep * (List1.ListIndex + 1), 10
Next
For e = val1 To (val1 - Int(10)) Step -1
Beep e * (List1.ListIndex + 1), 10
Next
Case Else
End Select
End Sub

Private Sub Form_Load()
List3.AddItem "Custom"
List3.AddItem "Old Computer"
List3.AddItem "Melody"
List3.AddItem "Zip"
List3.AddItem "Conformer"

List3.ListIndex = 0
List2.AddItem "None"
List2.AddItem "Key effect: Fade down hz"
List2.AddItem "Key effect: Fade up hz"
List2.AddItem "Key effect: Random Distort"
List2.AddItem "Key effect: Fixed Distort"
List1.AddItem 0
List1.AddItem 1
List1.AddItem 2
List1.AddItem 3
List1.AddItem 4
List1.AddItem 5
List1.AddItem 6
List1.AddItem 7
List1.ListIndex = 0
End Sub

Private Sub Ge_Click()
Beep 1, Text1.Text


End Sub

Private Sub List3_Click()
Select Case List3.ListIndex
Case 1
Text3.Text = 150
Case 2
Text3.Text = 45
Case 3
Text3.Text = 15
End Select
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If List3.ListIndex <> 4 Then
Beep X + Y, Text3.Text
Else
Beep X + Y, Text3.Text + (Y / 50)
End If
Label3.Caption = Hex(X) & " " & Hex(Y) & " - " & Hex(Text3.Text) & Oct(X + Y)
End Sub
