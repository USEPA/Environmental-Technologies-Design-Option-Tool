Attribute VB_Name = "Refresh"
Option Explicit




Const Refresh_declarations_end = True


Sub refresh_frmMain()
  
  Call AssignTextAndTag(frmMain.txtData(0), nowproj.dyestudy_calcdate)
  

End Sub

Sub refresh_frmDyeStudy(proj As Project_Type)
Dim i As Integer
  
  
  frmDyeStudy.f1book_dyestudy.Visible = False
  frmDyeStudy.f1book_dyestudy.MaxRow = proj.dyestudy_count
  For i = 1 To proj.dyestudy_count
    frmDyeStudy.f1book_dyestudy.EntryRC(i, 1) = _
        Trim$(Str$(proj.DyeStudy(i).time))
    frmDyeStudy.f1book_dyestudy.EntryRC(i, 2) = _
        Trim$(Str$(proj.DyeStudy(i).concentration))
  Next i
  frmDyeStudy.f1book_dyestudy.Visible = True
  Call AssignTextAndTag(frmDyeStudy.txtData(0), proj.dyestudy_calcdate)

End Sub


