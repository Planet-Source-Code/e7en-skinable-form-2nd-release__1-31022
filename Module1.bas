Attribute VB_Name = "Module1"
Dim ctrl(100)
Dim v
Dim Info

Sub SaveControls()
Dim Contrl As Control
For Each Contrl In Form1.Controls
typ = TypeName(Contrl)

Left1 = "Left=" & Contrl.Left
Top1 = "Top=" & Contrl.Top
Width1 = "Width=" & Contrl.Width
Height1 = "Height=" & Contrl.Height
value1 = ""

Select Case typ

Case "CommandBox"
value1 = Contrl.Caption

Case "Label"
value1 = Contrl.Caption

Case "TextBox"
value1 = Contrl.Text

Case "OptionButton"
value1 = Contrl.Value

Case "CheckBox"
value1 = Contrl.Value

End Select

If value1 = "" Then Info = Info & typ & "." & Contrl.Name & "." & Left1 & vbCrLf & typ & "." & Contrl.Name & "." & Top1 & vbCrLf & typ & "." & Contrl.Name & "." & Width1 & vbCrLf & typ & "." & Contrl.Name & "." & Height1 & vbCrLf & vbCrLf & vbCrLf Else Info = Info & typ & "." & Contrl.Name & "." & Left1 & vbCrLf & typ & "." & Contrl.Name & "." & Top1 & vbCrLf & typ & "." & Contrl.Name & "." & Width1 & vbCrLf & typ & "." & Contrl.Name & "." & Height1 & vbCrLf & typ & "." & Contrl.Name & "." & "Value=" & value1 & vbCrLf & vbCrLf

Next

Text1.Text = Info

Open "c:\temp.txt" For Output As #1
Print #1, Info
Close #1
End Sub

Sub LoadControls()
Dim str As String
Open "c:\temp.txt" For Input As #1
Do Until EOF(1)
Input #1, str
If str = "" Then Else AddValues Split(str, ".")(0), Split(str, ".")(1), Split(Split(str, "=")(0), ".")(2), Split(str, "=")(1)
Loop
Close #1
End Sub





Function AddValues(ByVal ControlType, ByVal ControlName, ByVal ControlValue As String, ByVal Value As String) As Boolean
Dim Contrl As Control
AddValues = False
For Each Contrl In Me.Controls

typ = TypeName(Contrl)

If typ = ControlType And Contrl.Name = ControlName Then
Select Case ControlValue

Case "Left"
Contrl.Left = Value

Case "Top"
Contrl.Top = Value

Case "Width"
Contrl.Width = Value

Case "Height"
Contrl.Height = Value

Case "Value"
Select Case typ

Case "CommandBox"
Contrl.Caption = Value

Case "Label"
Contrl.Caption = Value

Case "TextBox"
Contrl.Text = Value

Case "OptionButton"
Contrl.Value = Value

Case "CheckBox"
Contrl.Value = Value
End Select
End Select
AddValues = True
End If
Next
If AddValues = False Then
v = v + 1
temp = "vb." & ControlType
Set Contrl(v) = Me.Controls.Add(temp, ControlName, Me)
'ctrl(v).Name = ControlName
AddValues ControlType, ControlName, ControlValue, Value
End If
End Function

