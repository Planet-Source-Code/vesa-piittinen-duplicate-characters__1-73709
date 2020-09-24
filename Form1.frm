VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzA"
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ITERATIONS = 100000
Dim TESTSTRING As String

Dim M As HDM

Private Sub Command1_Click()
    Dim P As Long, I As Long
    Timing = 0
    For I = 1 To ITERATIONS
        P = HasDuplicates(TESTSTRING)
    Next I
    Command1.Caption = "HasDuplicates: " & Format$(Timing * 1000, "0.000000000\ \m\s") & " @ pos: " & P
End Sub

Private Sub Command2_Click()
    Dim P As Long, I As Long
    Timing = 0
    For I = 1 To ITERATIONS
        P = M.HasDuplicatesM(TESTSTRING)
    Next I
    Command2.Caption = "HasDuplicatesM: " & Format$(Timing * 1000, "0.000000000\ \m\s") & " @ pos: " & P
End Sub

Private Sub Form_Load()
    TESTSTRING = Text1.Text
    Set M = New HDM
    Debug.Print M.HasDuplicatesM(TESTSTRING)
    Debug.Assert Not Unloader
End Sub

Private Function Unloader() As Boolean
    Unload Me
End Function

Private Sub Text1_Change()
    TESTSTRING = Text1.Text
End Sub
