# VB


VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   7515
   ClientTop       =   3870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command17 
      Caption         =   "."
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "/"
      Height          =   375
      Left            =   2760
      TabIndex        =   23
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "*"
      Height          =   375
      Left            =   2760
      TabIndex        =   22
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "+"
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "="
      Height          =   855
      Left            =   3360
      TabIndex        =   19
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "AC"
      Height          =   855
      Left            =   3360
      TabIndex        =   18
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   1920
      TabIndex        =   17
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "First"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "当前状态"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "===="
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Frame2 = "First" Then
Text1 = Text1 + "1"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "1"
End If
End Sub

Private Sub Command10_Click()
If Frame2 = "First" Then
Text1 = Text1 + "0"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "0"
End If
End Sub

Private Sub Command11_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Frame1 = ""
Frame2 = "First"
Frame3 = ""
End Sub

Private Sub Command12_Click()
If Val(Text2) = 0 And Frame3 = "4" Then
Dim a As Integer
a = MsgBox("除数不能为0", 5 + 16 + 0 + 0, WRONG)
End If
If a = 4 Then
Text1 = ""
Text2 = ""
Text3 = ""
Frame1 = ""
Frame2 = "First"
Frame3 = ""
End If
If a = 2 Then
End
End If
If Frame3 = "1" Then
Text3 = Val(Text1) + Val(Text2)
End If
If Frame3 = "2" Then
Text3 = Val(Text1) - Val(Text2)
End If
If Frame3 = "3" Then
Text3 = Val(Text1) * Val(Text2)
End If
If Frame3 = "4" Then
Text3 = Val(Text1) / Val(Text2)
End If
End Sub

Private Sub Command13_Click()
Frame1 = "+"
Frame2 = "Second"
Frame3 = "1"
End Sub

Private Sub Command14_Click()
Frame1 = "-"
Frame2 = "Second"
Frame3 = "2"
End Sub

Private Sub Command15_Click()
Frame1 = "*"
Frame2 = "Second"
Frame3 = "3"
End Sub

Private Sub Command16_Click()
Frame1 = "/"
Frame2 = "Second"
Frame3 = "4"
End Sub

Private Sub Command17_Click()
If Frame2 = "First" Then
Text1 = Text1 + "."
End If
If Frame2 = "Second" Then
Text2 = Text2 + "."
End If
End Sub

Private Sub Command2_Click()
If Frame2 = "First" Then
Text1 = Text1 + "2"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "2"
End If
End Sub

Private Sub Command3_Click()
If Frame2 = "First" Then
Text1 = Text1 + "3"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "3"
End If
End Sub

Private Sub Command4_Click()
If Frame2 = "First" Then
Text1 = Text1 + "4"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "4"
End If
End Sub

Private Sub Command5_Click()
If Frame2 = "First" Then
Text1 = Text1 + "5"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "5"
End If
End Sub

Private Sub Command6_Click()
If Frame2 = "First" Then
Text1 = Text1 + "6"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "6"
End If
End Sub

Private Sub Command7_Click()
If Frame2 = "First" Then
Text1 = Text1 + "7"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "7"
End If
End Sub

Private Sub Command8_Click()
If Frame2 = "First" Then
Text1 = Text1 + "8"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "8"
End If
End Sub

Private Sub Command9_Click()
If Frame2 = "First" Then
Text1 = Text1 + "9"
End If
If Frame2 = "Second" Then
Text2 = Text2 + "9"
End If
End Sub

