VERSION 5.00
Begin VB.Form foRM6 
   AutoRedraw      =   -1  'True
   Caption         =   "Draw it"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3720
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   1800
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "1"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Draw it.frx":0000
         Left            =   1080
         List            =   "Draw it.frx":0002
         TabIndex        =   6
         Text            =   "13"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Draw Style:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Pen Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label4 
      Caption         =   "If you do not have a middle button use Left and right Button(s) = Blue"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Middle Button = Blue"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Left button = White"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Right Button = Red"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "foRM6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
On Error Resume Next
Me.DrawMode = Text1.Text
End Sub

Private Sub Form_Load()
Combo1.AddItem "0"
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
Combo1.AddItem "13"
Combo1.AddItem "14"
Combo1.AddItem "15"
Combo1.AddItem "16"
Combo2.AddItem "0"
Combo2.AddItem "1"
Combo2.AddItem "2"
Combo2.AddItem "3"
Combo2.AddItem "4"
Combo2.AddItem "5"
Combo2.AddItem "6"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    foRM6.CurrentX = X


        foRM6.CurrentY = Y
        End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Draws a White line if left button is clicked
    If Button = 1 Then
        Line (foRM6.CurrentX, foRM6.CurrentY)-(X, Y), vbWhite
    End If
    'Draws a Red line if right button is clicked
    If Button = 2 Then
        Line (foRM6.CurrentX, foRM6.CurrentY)-(X, Y), vbRed
    End If
    'Draws a Blue line if middle button is clicked
    If Button = 4 Then
        Line (foRM6.CurrentX, foRM6.CurrentY)-(X, Y), vbBlue
    End If
    'Draws a Blue line if left and right button is clicked
    If Button = 3 Then
        Line (foRM6.CurrentX, foRM6.CurrentY)-(X, Y), vbBlue
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    foRM6.CurrentX = X


        foRM6.CurrentY = Y
End Sub

Private Sub Text2_Change()
Me.DrawStyle = Text2.Text
End Sub

Private Sub Text3_Change()
On Error Resume Next
Me.DrawWidth = Text3.Text
End Sub

Private Sub Timer1_Timer()
Text1.Text = Combo1.Text
Text2.Text = Combo2.Text
End Sub
