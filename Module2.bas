Attribute VB_Name = "Module2"
Sub ClearData1()

'If Range("A4:C4").Value = "" Then
If WorksheetFunction.CountA(Range("A4:C4")) = 0 Then
    MsgBox ("Data is Blank")
Else
Run "ClearData2"

End If

End Sub

Sub ClearData2()

Dim answer As Integer

answer = MsgBox("Want to Clear all Data?", vbQuestion + vbYesNo + vbDefaultButton2, "WARNING")

If answer = vbYes Then

Cells(4, 1).Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete
End If

End Sub

