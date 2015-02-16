VERSION 5.00
Begin VB.Form frmComplicatedJoke 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complicated Joke"
   ClientHeight    =   2190
   ClientLeft      =   13005
   ClientTop       =   7230
   ClientWidth     =   3555
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3555
   Begin VB.CommandButton cmdPunchline 
      Caption         =   "Punchlin&e"
      Height          =   495
      Left            =   1110
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblComplicatedJoke 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What does the ""B"" in Benoit B. Mandelbrot stand for?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   330
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmComplicatedJoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPunchline_Click()
MsgBox "Benoit B. Mandelbrot!", None, "Geddit? Fractals?"
frmComplicatedJoke.Hide
frmMain.Show
End Sub
