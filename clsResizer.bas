Attribute VB_Name = "clsResizer"
Private mFormWidth As Long
Private mFormHeight As Variant
Private marrControlsSizes As Variant
Private miControlsCount As Integer
'Public Event ResizeAllComplete()

Public Sub GetMySizes(pForm As Form)
Dim X As Integer 'counter
On Error GoTo EH
'Set the object properties to the properties of the calling form...
mFormWidth = pForm.ScaleWidth
mFormHeight = pForm.ScaleHeight
miControlsCount = pForm.Controls.Count
'Size the array to hold the required number of controls by 6 (for the 6 properties to be stored for each ctl)
' ...these being a reference to the control itself and the top, left, width, height & font size of the control
ReDim marrControlsSizes(miControlsCount, 5)
'Populate the array...
For X = 0 To miControlsCount - 1
With pForm
marrControlsSizes(X, 0) = .Controls(X)
marrControlsSizes(X, 1) = .Controls(X).Top
marrControlsSizes(X, 2) = .Controls(X).Left
marrControlsSizes(X, 3) = .Controls(X).Width
marrControlsSizes(X, 4) = .Controls(X).Height
marrControlsSizes(X, 5) = .Controls(X).FontSize
End With
Next X

Exit Sub
EH:
If ERR.Number = 438 Or ERR.Number = 383 Then Resume Next
ERR.Raise vbObjectError + 600, "clsResizer.GetMySizes", ERR.Description
End Sub


Public Sub ResizeAll(pForm As Form)
Dim frm As Form
Dim Ctl As Control
Dim ScaleX As Double
Dim ScaleY As Double
Dim LeftArg As Double
On Error GoTo EH
'Set the object properties to the properties of the calling form...
Set frm = pForm
'Calculate the scale
ScaleX = frm.ScaleWidth / mFormWidth
ScaleY = frm.ScaleHeight / mFormHeight
'Iterate through the array of stored controls and original sizes to resize the controls on the form
For X = 0 To miControlsCount - 1
Set Ctl = frm.Controls(X)
With Ctl
'Set the font size...
If ScaleX < ScaleY Then
Ctl.FontSize = marrControlsSizes(X, 5) * ScaleX
Else
Ctl.FontSize = marrControlsSizes(X, 5) * ScaleY
End If
'The following is to enable compatibility with MS tabbed dialog control...
If marrControlsSizes(X, 2) < 0 Then
LeftArg = (marrControlsSizes(X, 2) + 75000) * ScaleX - 75000
Else
LeftArg = marrControlsSizes(X, 2) * ScaleX
End If
'=====================================================
Ctl.Move LeftArg, marrControlsSizes(X, 1) * ScaleY, marrControlsSizes(X, 3) * marrControlsSizes(X, 4) * ScaleY
Ctl.Width = marrControlsSizes(X, 3) * ScaleX
Ctl.Height = marrControlsSizes(X, 4) * ScaleY
End With
Next X
Set frm = Nothing
'RaiseEvent ResizeAllComplete
Exit Sub

EH:
If ERR.Number = 438 Or ERR.Number = 383 Then Resume Next
ERR.Raise vbObjectError + 601, "clsResizer.ResizeAll", ERR.Description
End Sub

