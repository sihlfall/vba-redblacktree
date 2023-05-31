VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton RunButton 
      Caption         =   "Run"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   4455
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Caption = "Press Button to run tests!"
End Sub

Private Sub RunButton_Click()
    Dim msg As String
    RedBlackTreeTest.RunTests msg
    Label1.Caption = msg & vbCrLf & "Tests completed successfully."
End Sub
