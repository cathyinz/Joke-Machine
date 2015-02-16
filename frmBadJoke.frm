VERSION 5.00
Begin VB.Form frmBadJoke 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bad Joke"
   ClientHeight    =   2415
   ClientLeft      =   12930
   ClientTop       =   7125
   ClientWidth     =   3555
   ControlBox      =   0   'False
   ForeColor       =   &H0080C0FF&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3555
   Begin VB.CommandButton cmdPunchline 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Punchlin&e"
      Height          =   495
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblBadJoke 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "An old lady at the bank asked me if I could help her check her balance."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   390
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmBadJoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPunchline_Click()
MsgBox "So I pushed her over.", None, "Oops"
frmBadJoke.Hide
frmMain.Show
End Sub

