Attribute VB_Name = "Print"
Option Explicit





Const Print_declarations_end = True



Sub Print_DyeStudy()
Dim NumCopies As Integer
Dim i As Integer
Dim j As Integer

  On Error GoTo err_ThisSub
  NumCopies = frmMain.CommonDialog1.Copies
  With nowproj
    For i = 1 To NumCopies
      If i > 1 Then Printer.NewPage
      Printer.FontName = "Times"
      Printer.FontSize = 14
      Printer.Print
      Printer.Print frmMain.Caption
      Printer.Print
      Printer.Print
      Printer.FontSize = 10
      Printer.Print "Last Calculated: " + nowproj.dyestudy_calcdate
      Printer.Print
      Printer.Print "Time"; Tab(30); "Concentration"
      For j = 1 To nowproj.dyestudy_count - 1
        Printer.Print Tab(2); nowproj.DyeStudy(j).time; Tab(32); nowproj.DyeStudy(j).concentration
      Next j
      Printer.Print
      Printer.Print
      Printer.Print nowproj.dyestudy_output
    Next i
    Printer.EndDoc
  End With
Dim msg As String
  msg = _
      "A total of " & _
      Trim$(Str$(NumCopies)) & _
      IIf(NumCopies = 1, " copy was ", " copies were ") & _
      "successfully printed."
  Call Show_Message(msg, vbExclamation, App.Title)
exit_normally_ThisSub:
  Exit Sub
exit_error_ThisSub:
  Exit Sub
err_ThisSub:
  Call Show_Trapped_Error("Print_Efficiency_Results")
  Resume exit_error_ThisSub
End Sub

