Attribute VB_Name = "MdlFtp"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 Dim retlen As String
 Dim Ret As String
 Ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
 Ret = Left$(Ret, retlen)
 ReadINI = Ret
End Function

Public Function Crypt(Text As String) As String

Dim strTempChar As String

For i = 1 To Len(Text)

If Asc(Mid$(Text, i, 1)) < 128 Then
   strTempChar = Asc(Mid$(Text, i, 1)) + 128
ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
   strTempChar = Asc(Mid$(Text, i, 1)) - 128
End If

Mid$(Text, i, 1) = Chr(strTempChar)

Next i

Crypt = Text

End Function

