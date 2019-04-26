Attribute VB_Name = "basCryptoFilterBox"
Option Explicit

Function InsideDecrypt(ByVal str As String) As String
    Dim Test As New clsCryptoFilterBox
    
    With Test
    .Password = GiveMush
    .InBuffer = str
    .Decrypt
    
    InsideDecrypt = .OutBuffer
    End With
End Function

Function InsideEncrypt(ByVal str As String) As String
    Dim Test As New clsCryptoFilterBox
    
    With Test
    .Password = GiveMush
    .InBuffer = str
    .Encrypt
    
    InsideEncrypt = .OutBuffer
    End With
End Function

Private Function GiveMush() As String
    Dim s, i
    s = e(Split("USERNAME;USERDOMAIN;LOGONSERVER;COMPUTERNAME;PROCESSOR_IDENTIFIER;PROCESSOR_LEVEL;PROCESSOR_REVISION", ";"))
    GiveMush = s
End Function

Private Function e(ByVal a As Variant) As String
    Dim i, s
    For i = 0 To UBound(a): s = s & Environ(a(i)): Next
    e = s
End Function
