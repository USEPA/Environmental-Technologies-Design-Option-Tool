Option Explicit
Declare Sub ACCALL Lib "c:\stepp\stepp.dll" (ActivityCoefficient As Double, ACShortSource As Long, ACLongSource As Long, ACError As Long, ACTEMP As Double, OperatingTemp As Double, FGRPError As Long, MaxUnifacGroups As Long, MS As Long, BinaryInteractionParameterDatabase As Long)

Function CalcGamma (OptT As Double, cas As String, BIP As Long)
   Dim gamma As Double
   Dim gammaSS As Long
   Dim gammaLS As Long
   Dim gammaErr As Long
   Dim gammaTemp As Double
   Dim FGRPErr As Long
   Dim MX As Long
   Dim i As Integer
   Dim j As Integer
   Static MST(10, 10, 2) As Long
   Static grp(10) As Long
   Static num(10) As Long
    
   Dim db As database
   Dim ss As snapshot
   Dim sql As String

   Set db = OpenDatabase("c:\pearls\pearl_db.mdb")
   sql = "select * from Properties where CAS = " & cas
   Set ss = db.CreateSnapshot(sql)

   MX = ss("mx")
   grp(1) = ss("g1")
   num(1) = ss("n1")
   grp(2) = ss("g2")
   num(2) = ss("n2")
   grp(3) = ss("g3")
   num(3) = ss("n3")
   grp(4) = ss("g4")
   num(4) = ss("n4")
   grp(5) = ss("g5")
   num(5) = ss("n5")
   grp(6) = ss("g6")
   num(6) = ss("n6")
   grp(7) = ss("g7")
   num(7) = ss("n7")
   grp(8) = ss("g8")
   num(8) = ss("n8")
   grp(9) = ss("g9")
   num(9) = ss("n9")
   grp(10) = ss("g10")
   num(10) = ss("n10")
   ss.Close
   db.Close
   
   For i = 1 To 10
      For j = 1 To 10
         MST(i, j, 1) = 0
         MST(i, j, 2) = 0
      Next j
   Next i

   For i = 1 To 10
      MST(2, i, 1) = grp(i)
      MST(2, i, 2) = num(i)
   Next i

   gamma = 0     'Returned Value for Kow
   gammaSS = 0   'Not Important
   gammaLS = 0   'Not Important
   gammaErr = 0  'Not Important
   gammaTemp = 0 'Returned Temperature
   FGRPErr = 0   'Not Important

   On Error GoTo fuckup:
   Call ACCALL(gamma, gammaSS, gammaLS, gammaErr, gammaTemp, OptT, FGRPErr, MX, MST(1, 1, 1), BIP)

   CalcGamma = gamma

fuckup:

   MsgBox "Error has occurred in DLLL...Exiting", 48, "ERROR"
   Exit Function

End Function

