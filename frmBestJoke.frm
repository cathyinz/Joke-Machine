VERSION 5.00
Begin VB.Form frmBestJoke 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Best Joke"
   ClientHeight    =   2280
   ClientLeft      =   12855
   ClientTop       =   7155
   ClientWidth     =   3555
   ControlBox      =   0   'False
   ForeColor       =   &H00C0FFC0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3555
   Begin VB.CommandButton cmdPunchline 
      Caption         =   """Knock knock!"""
      Height          =   495
      Left            =   1110
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblBestJoke 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I've got a great knock knock joke, but you've got to start it..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   570
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmBestJoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPunchline_Click()
MsgBox "Wait a minute", None, "Who's there?"
frmBestJoke.Hide
frmMain.Show
End Sub
