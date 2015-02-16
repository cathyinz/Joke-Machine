VERSION 5.00
Begin VB.Form frmOKJoke 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OK Joke"
   ClientHeight    =   1710
   ClientLeft      =   13005
   ClientTop       =   7425
   ClientWidth     =   3555
   ControlBox      =   0   'False
   ForeColor       =   &H00C0FFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3555
   Begin VB.CommandButton cmdPunchline 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Punchlin&e"
      Height          =   495
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblOKJoke 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What do you call a fake noodle?"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   450
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmOKJoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPunchline_Click()
MsgBox "An impasta!", None, "Mamma Mia!"
frmOKJoke.Hide
frmMain.Show
End Sub


