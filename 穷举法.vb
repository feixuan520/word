Private Sub Command1_Click()
Dim n As Long
n = Val(Text1.Text)

Dim i%, j%

For i = 2 To n Step 1 '穷举所有的数
    Dim flag As Boolean
    flag = True '默认是为素数
    
    For j = 2 To i - 1 Step 1 '对每个数进行穷举判断
        
        If i Mod j = 0 Then
            flag = False '不是素数
            Exit For
        End If
    Next j
    
    If flag Then Text2.Text = Text2.Text & i & ","

Next i
        
End Sub
