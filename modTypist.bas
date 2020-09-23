Attribute VB_Name = "modTypist"
Option Explicit


'Inspired by zebba's "A Ghost typer ( useing timers )" txtCodeId=63737.
'
'Assumes font properties are already set for the control
'ctrl As Variant allows various different controls to be used
'You could develop other routines to cope with the
'specific needs of other controls for displaying text
Public Sub Pause(ByVal dblInterval As Double)

  Dim EndTime As Double

  EndTime = dblInterval + Timer 'do the math once, makes the timer event more precise
  Do While Timer < EndTime
    DoEvents
  Loop

End Sub

Public Sub RandTypistCaption(ctrl As Variant, _
                             ByVal sMsg As String, _
                             Optional Speed As Long = 1)

'Caption controls should only be used for short messages as
'longer messages will flicker becuase the whole caption has to be rewritten

  Dim CPoint        As Long

  Do While CPoint <= Len(sMsg)
    CPoint = CPoint + 1
    ctrl.Caption = Left$(sMsg, CPoint)
    DoEvents
    Pause Rnd / Speed
  Loop

End Sub

Public Sub RandTypistText(ctrl As Variant, _
                          ByVal sMsg As String, _
                          Optional Speed As Long = 1, _
                          Optional ByVal bClearText As Boolean = True)

'creates a smoother effect for controls with the Text Property
'by using SelStart/SelText to append to the end of existing text
'Optional Parameter bClearText allows you to attach text to existing text if False
'By default this routine clears the text before displaying

  Dim CPoint        As Long

  If bClearText Then
    ctrl.Text = vbNullString
  End If
  Do While CPoint <= Len(sMsg)
    CPoint = CPoint + 1
    ctrl.SelStart = Len(ctrl.Text)
    ctrl.SelText = Mid$(sMsg, CPoint, 1)
    DoEvents
    Pause Rnd / Speed
  Loop

End Sub

':)Code Fixer V4.0.29 (Friday, 30 December 2005 14:07:46) 1 + 58 = 59 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 133302322223333232|33332222222222222222222222222202|1112222|2221222|222222222233|111111111111|1122222222220|333333|

