Attribute VB_Name = "Module2"
Public Sub LongSweep(Optional Ms As Byte)
If Ms <> 0 Then
GoTo sweep:
Else
Ms = 1
GoTo sweep:
End If
Exit Sub
sweep:
For eee = 0 To 3520
Beep eee, Ms
Next
End Sub
