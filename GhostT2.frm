VERSION 5.00
Begin VB.Form GhostT2 
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   Icon            =   "GhostT2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGhost 
      Caption         =   "Notes"
      Height          =   315
      Index           =   3
      Left            =   3600
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdGhost 
      Caption         =   "Details"
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdGhost 
      Caption         =   "Append"
      Height          =   315
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdGhost 
      Caption         =   "Intro"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtGhost 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblGhost 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   4335
   End
End
Attribute VB_Name = "GhostT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdGhost_Click(Index As Integer)

  Enabler cmdGhost, False
  Select Case Index
   Case 0 'Intro
    DoIntro
   Case 1 'Details
    RandTypistText txtGhost, "These routines take a control and a string message as parameters.", 70, True
    Pause 4
    RandTypistText txtGhost, vbNewLine & _
     "They also have an Optional parameter 'Speed'. This limits the maximum length of a pause between letters (one second/Speed). Values less than 5 produce the best illusion." & vbNewLine, 2, False
    Pause 6
    RandTypistText txtGhost, vbNewLine & _
     "The RandTypistText routine has an additional Optional parameter 'bClearText' which allows you to clear the previous text first (Default True). ", 90, False
    RandTypistText txtGhost, vbNewLine & _
     "Click 'Append' the effect of setting this parameter to False.", 100, False
   Case 2 'Append
    RandTypistText txtGhost, vbNewLine & _
     "RandTypistText can optionally append new text to an existing text." & vbNewLine & "Click 'Notes' to see the effect of 'bClearText' = True", 60, False
   Case 3 'Notes
    RandTypistText txtGhost, "Because the Caption property has to be completely redisplayed you should not use RandTypistCaption to display long messages as they will flicker while updating. Watch the label below as a long message is feed to it", 50
    RandTypistCaption lblGhost, "Demo of the flicker effect for very long caption messages", 80
    RandTypistText txtGhost, vbNewLine & _
     "RandTypistText can append to existing text so does not flicker(much).", 10, False
    RandTypistText txtGhost, vbNewLine & _
     "These routines do not attempt to set the Font or colour of the control. You might want to add these complications.", 10, False
    RandTypistCaption lblGhost, "ENJOY!"
  End Select
  Enabler cmdGhost, True

End Sub

Private Sub DoIntro()

'allows 2 different routines to call the intro
 cmdGhost(0).Caption = ""
cmdGhost(1).Caption = ""
cmdGhost(2).Caption = ""
cmdGhost(3).Caption = ""
lblGhost.Caption = ""
Me.Caption = ""
  RandTypistCaption Me, "Ghost Typer 2", 30
  RandTypistText txtGhost, "Sub-routines make coding simpler and more flexible!" & vbNewLine & _
   "This code uses 2 sub-routines that allow you to type messages to either the Caption or Text property of a control without using Timer controls.", 10
  RandTypistCaption lblGhost, "ENJOY!"
  RandTypistText txtGhost, vbNewLine & "Click 'Details' for more.", 110, False
  
RandTypistCaption cmdGhost(0), "Intro", 4
RandTypistCaption cmdGhost(1), "Details", 3
RandTypistCaption cmdGhost(2), "Append", 2
RandTypistCaption cmdGhost(3), "Notes"
Enabler cmdGhost, True
End Sub

Private Sub Enabler(CtrlArray As Variant, _
                    ByVal bEnabled As Boolean)

  Dim C As Variant

  For Each C In CtrlArray
    C.Enabled = bEnabled
  Next C

End Sub

Private Sub Form_Load()

  Show 'force form to show before the DoIntro routine fires
  Randomize
  Enabler cmdGhost, False
  DoIntro

End Sub

Private Sub Form_Unload(Cancel As Integer)

  End

End Sub

':)Code Fixer V4.0.29 (Friday, 30 December 2005 14:07:46) 7 + 66 = 73 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 133302322223333232|33332222222222222222222222222202|1112222|2221222|222222222233|111111111111|1122222222220|333333|

