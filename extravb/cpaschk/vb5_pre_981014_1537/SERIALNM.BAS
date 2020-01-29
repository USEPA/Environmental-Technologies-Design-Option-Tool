Attribute VB_Name = "SERIALNM"
Option Explicit

Const SNUMS_KEY1 = 230123967
Const SNUMS_KEY2 = 758923042
Const SNUMS_MAX = 1000000





Const SERIALNM_declarations_end = True


Function Snums_Generate(xin As Long) As String
Dim xout1 As Long
Dim xout2 As Long
Dim xout3 As Long
Dim xout5 As Long
Dim xkey1 As Long
Dim xkey2 As Long
Dim sout1 As String
Dim sout2 As String
Dim sout3 As String
Dim sout5 As String
Dim s As String
Dim keyidx1() As Integer
Dim keyidx2() As Integer
Dim keyidx3() As Integer
Dim keyidx4() As Integer
Dim keyidx5() As Integer
Dim i As Integer
  xkey1 = SNUMS_KEY1
  xkey2 = SNUMS_KEY2
  xout1 = xin Xor xkey1
  xout2 = xin + xkey2
  Randomize
  xout3 = CLng(Rnd * 1000000000#)
  xout5 = CLng(Rnd * 10#)
        'MsgBox xout5
  sout1 = Trim$(Str$(xout1))
  sout2 = Trim$(Str$(xout2))
  sout3 = Trim$(Str$(xout3))
  sout5 = Trim$(Str$(xout5))
  'NOTE: xout5 IS THE NUMBER OF PLACES TO SHIFT THE INDEXES TO THE RIGHT.
  Call Snums_GetKeyIndexes(keyidx1(), keyidx2(), keyidx3(), keyidx4(), keyidx5())
  Call Snums_ShiftKeyIndexRight(keyidx1(), xout5)
  Call Snums_ShiftKeyIndexRight(keyidx2(), xout5)
  Call Snums_ShiftKeyIndexRight(keyidx3(), xout5)
  Call Snums_ShiftKeyIndexRight(keyidx4(), xout5)
  s = String$(6 + 1 + 10 + 1 + 10, "x")
  For i = 1 To UBound(keyidx1)
    Mid$(s, keyidx1(i), 1) = Mid$(sout1, i, 1)
  Next i
  For i = 1 To UBound(keyidx2)
    Mid$(s, keyidx2(i), 1) = Mid$(sout2, i, 1)
  Next i
  For i = 1 To UBound(keyidx3)
    Mid$(s, keyidx3(i), 1) = Mid$(sout3, i, 1)    '"."
  Next i
  For i = 1 To UBound(keyidx4)
    Mid$(s, keyidx4(i), 1) = "-"
  Next i
  For i = 1 To UBound(keyidx5)
    Mid$(s, keyidx5(i), 1) = Left$(sout5, 1)
  Next i
  Snums_Generate = s
End Function


Sub Snums_GetKeyIndexes(keyidx1() As Integer, keyidx2() As Integer, keyidx3() As Integer, keyidx4() As Integer, keyidx5() As Integer)
  ReDim keyidx1(1 To 9)
  ReDim keyidx2(1 To 9)
  ReDim keyidx3(1 To 7)
  ReDim keyidx4(1 To 2)
  ReDim keyidx5(1 To 1)
  keyidx1(1) = 11
  keyidx1(2) = 23
  keyidx1(3) = 16
  keyidx1(4) = 4
  keyidx1(5) = 27
  keyidx1(6) = 17
  keyidx1(7) = 19
  keyidx1(8) = 6
  keyidx1(9) = 13
  keyidx2(1) = 26
  keyidx2(2) = 2
  keyidx2(3) = 8
  keyidx2(4) = 12
  keyidx2(5) = 9
  keyidx2(6) = 15
  keyidx2(7) = 21
  keyidx2(8) = 28
  keyidx2(9) = 22
  keyidx3(1) = 1
  keyidx3(2) = 3
  keyidx3(3) = 5
  keyidx3(4) = 10
  'keyidx3(5) = 14
  keyidx3(5) = 20
  keyidx3(6) = 24
  keyidx3(7) = 25
  keyidx4(1) = 7
  keyidx4(2) = 18
  keyidx5(1) = 14
End Sub


Sub Snums_ShiftKeyIndexLeft(keyidx() As Integer, ShiftBy As Long)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim xold() As Integer
  ReDim xold(1 To UBound(keyidx))
  For i = 1 To ShiftBy
    For j = 1 To UBound(keyidx)
      xold(j) = keyidx(j)
    Next j
    For j = 1 To UBound(keyidx)
      k = j + 1
      If (k > UBound(keyidx)) Then k = 1
      keyidx(j) = xold(k)
    Next j
  Next i
End Sub


Sub Snums_ShiftKeyIndexRight(keyidx() As Integer, ShiftBy As Long)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim xold() As Integer
  ReDim xold(1 To UBound(keyidx))
  For i = 1 To ShiftBy
    For j = 1 To UBound(keyidx)
      xold(j) = keyidx(j)
    Next j
    For j = 1 To UBound(keyidx)
      k = j - 1
      If (k < 1) Then k = UBound(keyidx)
      keyidx(j) = xold(k)
    Next j
  Next i
End Sub


'RETURNS:
'  TRUE = PASSES THE TEST.
'  FALSE = INVALID SERIAL NUMBER.
Function Snums_Verify(s As String) As Integer
Dim xkey1 As Long
Dim xkey2 As Long
Dim xout1 As Long
Dim xout2 As Long
Dim xout5 As Long
Dim xoutA As Long
Dim xoutB As Long
Dim sout1 As String
Dim sout2 As String
Dim sout3 As String
Dim sout4 As String
Dim sout5 As String
Dim keyidx1() As Integer
Dim keyidx2() As Integer
Dim keyidx3() As Integer
Dim keyidx4() As Integer
Dim keyidx5() As Integer
Dim i As Integer
  xkey1 = SNUMS_KEY1
  xkey2 = SNUMS_KEY2
  If (Len(s) <> 6 + 1 + 10 + 1 + 10) Then
    Snums_Verify = False
    Exit Function
  End If
  Call Snums_GetKeyIndexes(keyidx1(), keyidx2(), keyidx3(), keyidx4(), keyidx5())
  sout5 = Mid$(s, keyidx5(1), 1)
  xout5 = CLng(sout5)
        'MsgBox xout5
  Call Snums_ShiftKeyIndexRight(keyidx1(), xout5)
  Call Snums_ShiftKeyIndexRight(keyidx2(), xout5)
  Call Snums_ShiftKeyIndexRight(keyidx3(), xout5)
  Call Snums_ShiftKeyIndexRight(keyidx4(), xout5)
  sout1 = String$(UBound(keyidx1), "x")
  sout2 = String$(UBound(keyidx2), "x")
  sout3 = String$(UBound(keyidx3), "x")
  sout4 = String$(UBound(keyidx4), "x")
  For i = 1 To UBound(keyidx1)
    Mid$(sout1, i, 1) = Mid$(s, keyidx1(i), 1)
  Next i
  For i = 1 To UBound(keyidx2)
    Mid$(sout2, i, 1) = Mid$(s, keyidx2(i), 1)
  Next i
  For i = 1 To UBound(keyidx3)
    Mid$(sout3, i, 1) = Mid$(s, keyidx3(i), 1)
  Next i
  For i = 1 To UBound(keyidx4)
    Mid$(sout4, i, 1) = Mid$(s, keyidx4(i), 1)
  Next i
  xout1 = CLng(sout1)
  xout2 = CLng(sout2)
  xoutA = xout1 Xor xkey1
  xoutB = xout2 - xkey2
  If (xoutA = xoutB) Then
    If (xoutA > SNUMS_MAX) Then
      Snums_Verify = False
      Exit Function
    End If
    Snums_Verify = True
    Exit Function
  Else
    Snums_Verify = False
    Exit Function
  End If

End Function

