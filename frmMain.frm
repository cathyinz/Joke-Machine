VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Joke Machine"
   ClientHeight    =   4005
   ClientLeft      =   12255
   ClientTop       =   6225
   ClientWidth     =   5025
   ControlBox      =   0   'False
   ForeColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5025
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000080FF&
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4125
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1035
      Width           =   615
   End
   Begin VB.CommandButton cmdComplicatedJoke 
      BackColor       =   &H00FFFF80&
      Caption         =   "Complicated Joke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2475
      Width           =   3600
   End
   Begin VB.CommandButton cmdBadJoke 
      BackColor       =   &H00FFFF80&
      Caption         =   "Bad Joke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1035
      Width           =   3600
   End
   Begin VB.CommandButton cmdOKJoke 
      BackColor       =   &H00FFFF80&
      Caption         =   "OK Joke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1755
      Width           =   3600
   End
   Begin VB.CommandButton cmdBestJoke 
      BackColor       =   &H00FFFF80&
      Caption         =   "Best Joke"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3195
      Width           =   3600
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pick a joke - ANY joke!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   285
      TabIndex        =   0
      Top             =   315
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBestJoke_Click()
frmMain.Hide
frmBestJoke.Show
End Sub

Private Sub cmdExit_Click()
MsgBox "Thanks for using my program!", None, "Goodbyeee"
End
End Sub

Private Sub cmdOKJoke_Click()
frmMain.Hide
frmOKJoke.Show
End Sub

Private Sub cmdBadJoke_Click()
frmMain.Hide
frmBadJoke.Show
End Sub

Private Sub cmdComplicatedJoke_Click()
frmMain.Hide
frmComplicatedJoke.Show
End Sub
