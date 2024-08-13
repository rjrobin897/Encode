' -------------------------------------------------
'  Simple Encryption (Xor)
' -------------------------------------------------
'
'

Sub Main()

  Dim i
  Dim m  
  Dim k

  Dim strMessage
  Dim strPassword

  Dim strChar
  Dim strResult

  Dim strMsg

  strMessage = "Message" ' Message Digest
  strPassword = "Password123" ' Secret Password

  For i = 1 to Len(strMessage)

    If m = Len(strPassword) Then
      m = 1
    Else
      m = (m + 1)
    End If

    For k = 1 to Len(strPassword)

      strChar = Chr(Asc(Mid(strMessage,i,1)) Xor Asc(Mid(strPassword,m,1)) Mod Asc(Chr(m)))
      strResult = strResult & strChar

      Exit For

    Next

  Next

  strMsg = InputBox("Result: ","Notice",strResult)

End Sub

Main()
