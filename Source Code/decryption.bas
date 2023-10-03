Attribute VB_Name = "decryption"
Public Function Decrypt(ByVal pwd As String) As String
 Dim Length As Integer
 Dim NewText As String
 Dim Charactor As String
 Dim i As Integer
 
 Charactor = ""
    Length = Len(pwd)
    For i = 1 To Length
        Charactor = Mid(pwd, i, 1)
        Select Case Asc(Charactor)
            Case 192 To 217     ' À - Ù ( A - Z )
                Charactor = Chr(Asc(Charactor) - 127)
            Case 218 To 243     ' Ú - ó ( a - z )
                Charactor = Chr(Asc(Charactor) - 121)
            Case 244 To 253     ' ô - ý ( 0 - 9 )
                Charactor = Chr(Asc(Charactor) - 196)
            Case 32             ' Space
                Charactor = Chr(32)
        End Select
        NewText = NewText + Charactor
    Next
    Decrypt = NewText
End Function

