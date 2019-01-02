Option Explicit
' Require SHA256.bas

' The MIT License (MIT)
' Copyright (c) 2018 Oscar Casas
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

Function Merkle(hashList() As String) As String
    Dim i As Long
    Dim NewList() As String
    ReDim NewList(0)
    If UBound(hashList) = 1 Then
        Merkle = hashList(1)
        Exit Function
    End If
    For i = 1 To UBound(hashList) - 1 Step 2
        ReDim Preserve NewList(UBound(NewList) + 1)
        NewList(UBound(NewList)) = HASH2(hashList(i), hashList(i + 1))
        If (UBound(hashList) Mod 2) = 1 Then
            ReDim Preserve NewList(UBound(NewList) + 1)
            NewList(UBound(NewList)) = HASH2(hashList(UBound(hashList)), hashList(UBound(hashList)))
        End If
        Merkle = Merkle(NewList)
    Next
End Function

Function HASH2(a As String, b As String) As String
    Dim a1() As Byte
    Dim b1() As Byte
    a1 = BinReverse(HexToBin(a))
    b1 = BinReverse(HexToBin(b))
    HASH2 = BinToHex(BinReverse(BINSHA256(BINSHA256(Cat32Bit(a1, b1)))))
End Function

Function BinReverse(b() As Byte) As Byte()
    Dim i As Long
    Dim Ret() As Byte
    Dim u As Long
    u = UBound(b)
    ReDim Ret(u)
    For i = u To 0 Step -1
        Ret(i) = b(u - i)
    Next
    BinReverse = Ret
End Function

Function HexToBin(s As String) As Byte()
    Dim Ret() As Byte
    Dim i As Long
    Dim u As Long
    Dim n As Long
    'Debug.Print s
    
    u = Len(s)
    ReDim Ret(u / 2 - 1)
    For i = 1 To u Step 2
        Ret(n) = CLng("&H" + Mid$(s, i, 2))
        n = n + 1
    Next
    HexToBin = Ret
End Function

Function BinToHex(b() As Byte) As String
    Dim i As Long
    Dim Ret As String
    For i = 0 To UBound(b)
        Ret = Ret + Right$("00" + Hex$(b(i)), 2)
    Next
    BinToHex = LCase$(Ret)
End Function

Function BinToStr(b() As Byte) As String
    Dim i As Long
    Dim Ret As String
    For i = 0 To UBound(b)
        Ret = Ret + Chr$(b(i))
    Next
    BinToStr = Ret
End Function

Function Cat32Bit(b1() As Byte, b2() As Byte) As Byte()
    Dim Ret() As Byte
    Dim i As Long
    Dim u1 As Long
    Dim u2 As Long
    u1 = UBound(b1)
    u2 = UBound(b2)
    ReDim Ret(u1 + u2 + 1)
    For i = 0 To u1
        Ret(i) = b1(i)
    Next
    For i = 0 To u2
        Ret(i + u1 + 1) = b2(i)
    Next
    Cat32Bit = Ret
End Function
