Sub ファイル名見直し()
    Dim mypath, MyFile$, m As Integer
    For pathIndex = 1 To 100
        mypath = Range("F" & pathIndex).Value
        If (mypath = "") Then
         Exit For
        End If
        If dir(mypath & "*", vbDirectory) = "" Then
            MsgBox "パス不存在"
            Exit Sub
        End If
        
        If Range("b65536").End(xlUp).Row < Range("a65536").End(xlUp).Row Then
            MsgBox "新しいファイル名を指定してください"
            Exit Sub
        End If
        mypath = mypath & "\"
        arr = Range("A2:B" & Range("A65536").End(xlUp).Row)
        For i = 1 To UBound(arr)
            If dir(mypath & arr(i, 1)) <> "" Then
                Name mypath & arr(i, 1) As mypath & arr(i, 2) & ".html"
            End If
       Next
    Next

    MsgBox "*****ファイル名を見直しました^.^ *****"
End Sub
   
Sub getFolderFiles(sFolderPath As String)
On Error Resume Next
Dim f As String
Dim file() As String
Dim x
k = 1
x = 2
ReDim file(1)
file(1) = sFolderPath & "\"
  
    f = Dir(file(1) & "*.html")
    Do Until f = ""
       Range("F" & x).Hyperlinks.Add Anchor:=Range("F" & x), Address:=file(i) & f, TextToDisplay:=f
        x = x + 1
        f = Dir
    Loop
  
End Sub
