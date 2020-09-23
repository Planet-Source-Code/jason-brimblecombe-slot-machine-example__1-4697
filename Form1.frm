VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SLOT MACHINE"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click here!"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Image p3 
      Height          =   495
      Left            =   5760
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   615
   End
   Begin VB.Image p2 
      Height          =   615
      Left            =   3120
      Picture         =   "Form1.frx":55E6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Image p1 
      Height          =   615
      Left            =   0
      Picture         =   "Form1.frx":ABCC
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Label LBLMONEY 
      BackStyle       =   0  'Transparent
      Caption         =   "$MONEY"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":101B2
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   5895
   End
   Begin VB.Image lever 
      Height          =   2175
      Left            =   4440
      Picture         =   "Form1.frx":10274
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image slot3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   3360
      Stretch         =   -1  'True
      ToolTipText     =   "http://get.to/JasonsVBPage"
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image slot2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   2040
      Stretch         =   -1  'True
      ToolTipText     =   "Visit Jason's VB Page sometime!"
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image slot1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   720
      Stretch         =   -1  'True
      ToolTipText     =   "This Example was downloaded from Jason's VB Page! http://get.to/JasonsVBPage"
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example has been downloaded from Jason's VB Page!
    'http://get.to/jasonsvbpage
' Thankyou for visiting!
Option Explicit
Dim Money As String
Dim N1 As String
Dim N2 As String
Dim N3 As String

Private Sub Command1_Click()
MsgBox "Every time you pull the lever, you lose $100 and if you win, you get $200!", vbInformation, "Did you know..."
MsgBox "Email the creator: jbrimble@hotmail.com", vbExclamation, "Email me!"
End Sub

Private Sub Form_Load()
'Give the user $1000 to start off with and show the amount of money on the label: LBLMONEY
Money = 1000
LBLMONEY.Caption = "$" + Money
' Show the Messagebox at startup
MsgBox "Welcome to the SLOTMACHINE Example!", vbInformation, "Hi!"
End Sub


Private Sub lever_Click()
Money = Money - 100
LBLMONEY.Caption = "$" + Money
'random images
N1 = Int(Rnd * 10)
N2 = Int(Rnd * 10)
N3 = Int(Rnd * 10)
'1st slot
If N1 = 1 Or N1 < 1 Then
slot1.Picture = p1.Picture
End If
If N1 = 2 Then
slot1.Picture = p2.Picture
End If
If N1 = 3 Or N1 > 3 Then
slot1.Picture = p3.Picture
End If
'2nd slot
If N2 = 1 Or N1 < 1 Then
slot2.Picture = p1.Picture
End If
If N2 = 2 Then
slot2.Picture = p2.Picture
End If
If N2 = 3 Or N1 > 3 Then
slot2.Picture = p3.Picture
End If
'3rd slot
If N3 = 1 Or N1 < 1 Then
slot3.Picture = p1.Picture
End If
If N3 = 2 Then
slot3.Picture = p2.Picture
End If
If N3 = 3 Or N1 > 3 Then
slot3.Picture = p3.Picture
End If
'now to check if the user has won!!!
'(Just be grateful that you don't have to type all of this out!)
If slot1.Picture = slot2.Picture And slot1.Picture = slot3.Picture Then
MsgBox "You have won $200!", vbExclamation, "You're a winner!"
Money = Money + 200
'refresh the money label AGAIN!!!
LBLMONEY.Caption = "$" + Money
End If
If Money > 5000 Then
End If
End Sub
