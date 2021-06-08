Attribute VB_Name = "LICFILE"
Option Explicit

Const LICFILE_KEY = 192
Const LICFILE_STRINGSIZE_MAX = 90
'NOTE: LICFILE_STRINGSIZE_MAX CAN'T BE INCREASED PAST ABOUT 254 UNLESS
'THE ENCRYPT/DECRYPT ROUTINES ARE REVISED.

Global LicFile_keyidx() As Integer

Type LicFileString_Type
  StrSize As String * 3
  TheStr As String * LICFILE_STRINGSIZE_MAX
End Type





Const LICFILE_declarations_end = True


Sub LicFile_AppendGarbage(fn As String, newcharsize As Long)
Dim f As Integer
Dim temp() As Integer
Dim i As Integer
Dim n As Integer
  n = newcharsize / 2
  ReDim temp(1 To n) As Integer
  Randomize
  For i = 1 To n
    temp(i) = CInt((Rnd * 65520#) - 32766)
  Next i
  f = FreeFile
  Open fn For Binary Access Read Write As #f
  Put #f, LOF(f) + 1, temp(1)
  For i = 1 To n
    Put #f, , temp(i)
  Next i
  Close #f
  Erase temp
End Sub


Sub LicFile_Create(fn As String, maxsize As Long)
Dim f As Integer
Dim temp() As Integer
Dim i As Integer
Dim n As Integer
  n = maxsize / 2
  ReDim temp(1 To n) As Integer
  Randomize
  For i = 1 To n
    temp(i) = CInt((Rnd * 65520#) - 32766)
  Next i
  f = FreeFile
  Open fn For Binary Access Read Write As #f
  For i = 1 To n
    Put #f, , temp(i)
  Next i
  Close #f
  Erase temp
End Sub


Function LicFile_DecryptString(s As String) As String
Dim i As Integer
Dim outStr As String
Dim thisc As Integer
  outStr = ""
  For i = 1 To Len(s)
    thisc = Asc(Mid$(s, i, 1))
    outStr = outStr & Chr$(LicFile_keyidx(1, thisc Xor i))
  Next i
  LicFile_DecryptString = outStr
End Function


Function LicFile_EncryptString(s As String) As String
Dim i As Integer
Dim outStr As String
Dim thisc As Integer
  outStr = ""
  For i = 1 To Len(s)
    thisc = Asc(Mid$(s, i, 1))
    outStr = outStr & Chr$(LicFile_keyidx(0, thisc) Xor i)
  Next i
  LicFile_EncryptString = outStr
End Function


Sub LicFile_GenerateEncryptionKey()
Dim i As Integer
  ReDim LicFile_keyidx(0 To 1, 0 To 255)
  For i = 0 To 255
    LicFile_keyidx(0, i) = (i * 3) Mod 256
    LicFile_keyidx(1, (i * 3) Mod 256) = i
  Next i
  'STRICTLY SPEAKING, THIS ISN'T REALLY ENCRYPTION; IT'S MORE LIKE
  'A JOKE.  AS THEY SAY, "LOCKS ONLY KEEP HONEST PEOPLE OUT."
  'THE EASIEST WAY TO BREAK THROUGH A SERIAL NUMBER CHECK IS TO MODIFY
  'THE PROGRAM SO IT DOESN'T PERFORM THE CHECK.  THAT WOULD MAKE
  'ANY ENCRYPTION A MOOT POINT ANYWAY.
End Sub


Sub LicFile_GetEncryptedString(fn As String, pos As Long, retstr As String)
Dim f As Integer
Dim lfs As LicFileString_Type
Dim UseStr As String
Dim UseStrSize As Integer
  f = FreeFile
  Open fn For Binary Access Read Write As #f
  Get #f, pos, lfs
  Close #f
  lfs.StrSize = LicFile_DecryptString(lfs.StrSize)
  UseStrSize = Val(lfs.StrSize)
  If (UseStrSize > 0) Then
    UseStr = Left$(lfs.TheStr, UseStrSize)
  Else
    UseStr = ""
  End If
  UseStr = LicFile_DecryptString(UseStr)
  retstr = UseStr
End Sub


Sub LicFile_PutEncryptedString(fn As String, pos As Long, s As String)
Dim f As Integer
Dim lfs As LicFileString_Type
Dim UseStr As String
Dim UseStrSize As Integer
Dim i As Integer
  f = FreeFile
  If (Len(s) > LICFILE_STRINGSIZE_MAX) Then
    UseStr = Left$(s, LICFILE_STRINGSIZE_MAX)
  Else
    UseStr = s
  End If
  UseStr = LicFile_EncryptString(UseStr)
  UseStrSize = Len(UseStr)
  Randomize
  For i = UseStrSize + 1 To LICFILE_STRINGSIZE_MAX
    UseStr = UseStr & Chr$(Int(Rnd * 255))
  Next i
  Open fn For Binary Access Read Write As #f
  lfs.StrSize = Trim$(Str$(Len(s)))
  Do While (Len(lfs.StrSize) < 3)
    lfs.StrSize = lfs.StrSize & " "
  Loop
  lfs.StrSize = LicFile_EncryptString(lfs.StrSize)
  lfs.TheStr = UseStr
  Put #f, pos, lfs
  Close #f
End Sub


Sub Parser_GetArg(sepchar As String, inline As String, ArgNum As Integer, retstr As String)
Dim i As Integer
Dim j As Integer
  retstr = ""
  j = 1
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      j = j + 1
      If (j > ArgNum) Then Exit For
    Else
      If (j = ArgNum) Then
        retstr = retstr + Mid$(inline, i, 1)
      End If
    End If
  Next i
End Sub


Function Parser_GetNumArgs(sepchar As String, inline As String) As Integer
Dim NumArgs As Integer
Dim i As Integer
  NumArgs = 1     'between chr #1 and first separator char.
  For i = 1 To Len(inline)
    If (Mid$(inline, i, 1) = sepchar) Then
      NumArgs = NumArgs + 1
    End If
  Next i
  Parser_GetNumArgs = NumArgs
End Function


Function Parser_RemoveCharacters(remove_char As String, inline As String) As String
Dim retstr As String
Dim i As Integer
Dim ok_append As Integer
Dim thisc As String
  retstr = ""
  For i = 1 To Len(inline)
    ok_append = True
    thisc = Mid$(inline, i, 1)
    If (thisc = remove_char) Then ok_append = False
    If (ok_append) Then
      retstr = retstr & thisc
    End If
  Next i
  Parser_RemoveCharacters = retstr
End Function


Function Parser_RemoveDuplicateSeparators(sepchar As String, inline As String) As String
Dim retstr As String
Dim i As Integer
Dim ok_append As Integer
Dim thisc As String
  retstr = ""
  For i = 1 To Len(inline)
    ok_append = True
    thisc = Mid$(inline, i, 1)
    If (i > 1) Then
      If (thisc = sepchar) Then
        If (Right$(retstr, 1) = sepchar) Then
          ok_append = False
        End If
      End If
    End If
    If (ok_append) Then
      retstr = retstr & thisc
    End If
  Next i
  Parser_RemoveDuplicateSeparators = retstr
End Function

