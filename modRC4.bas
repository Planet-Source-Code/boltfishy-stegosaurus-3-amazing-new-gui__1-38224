Attribute VB_Name = "RC4"

Public Function cRC4(inp As String, key As String) As String
On Error Resume Next

Dim s(0 To 255) As Byte, K(0 To 255) As Byte, i As Long
Dim j As Long, temp As Byte, y As Byte, t As Long, x As Long
Dim Outp As String

For i = 0 To 255
s(i) = i
Next

j = 1
For i = 0 To 255
If j > Len(key) Then j = 1
K(i) = Asc(Mid(key, j, 1))
j = j + 1
Next i

j = 0
For i = 0 To 255
j = (j + s(i) + K(i)) Mod 256
temp = s(i)
s(i) = s(j)
s(j) = temp
Next i

i = 0
j = 0
For x = 1 To Len(inp)
i = (i + 1) Mod 256
j = (j + s(i)) Mod 256
temp = s(i)
s(i) = s(j)
s(j) = temp
t = (s(i) + (s(j) Mod 256)) Mod 256
y = s(t)
    
Outp = Outp & Chr(Asc(Mid(inp, x, 1)) Xor y)
Next
cRC4 = Outp

End Function

