Attribute VB_Name = "ModSupportMenu"
Option Explicit

Public Sub MenuChecking(Ctrl As Variant, _
                        ByVal chekMe As Long)

'generic for controlling Checks in menu arrays
'unchecks all except the selected member
'requires continuous menu array

  Dim I As Long

  For I = 0 To Ctrl.Count - 1
    Ctrl(I).Checked = I = chekMe
  Next I

End Sub

':)Code Fixer V2.8.3 (6/01/2005 7:58:15 AM) 1 + 17 = 18 Lines Thanks Ulli for inspiration and lots of code.

