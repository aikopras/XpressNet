Attribute VB_Name = "ModFunctions"
Public Sub CheckFOnOff(ByVal index As Integer)
  FormConnect.LabelStatus = "Klopt"
   Debug.Print "index = " & index
       FormConnect.LEDF0F68(3).FillColor = vbRed
    FormConnect.LEDF0F68(5).FillColor = vbGreen
    
    FormConnect.LEDF0F68(4).FillColor = vbGreen
    
  If index < 0 Then Exit Sub
  If index > 68 Then Exit Sub
  If FormConnect.CheckFOnOff.value Then
    FormConnect.LEDF0F68(index).FillColor = vbRed
  Else
    FormConnect.LEDF0F68(index).FillColor = vbGreen
  End If
  DoEvents
End Sub

