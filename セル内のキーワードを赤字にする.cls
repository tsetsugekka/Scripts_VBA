Option Explicit 
 
Sub Mark() 
    Dim rng As Range 
    Dim str As String 
    Dim s As Integer 
    str = "キーワード" 
     
    For Each rng In Selection 
        If InStr(1, rng.Value, str) > 0 Then 
            s = InStr(1, rng.Value, str) 
            With rng.Characters(s, Len(str)).Font 
                .Color = vbRed 
            End With 
        End If 
    Next rng 
End Sub 
