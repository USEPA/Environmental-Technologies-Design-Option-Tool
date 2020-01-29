Attribute VB_Name = "modunits"
Attribute VB_Description = "String*"
Option Explicit

Sub ConvertTFTUnits(ConvertFrom As String, ConvertTo As String)
    
    Dim i As Integer
    Dim J As Integer
    Dim TempProp As Integer
    Dim DefaultUnit As String
    
    If Trim(ConvertFrom) = Trim(ConvertTo) Then Exit Sub
                
    DefaultUnit = Get_DefaultUnit(-1)
    
    If Trim(ConvertFrom) <> Trim(DefaultUnit) Then
        For J = 0 To NumProperties
            TempProp = CurProp
            CurProp = J
            InfoMethod(CurProp).TFT = Convert(InfoMethod(CurProp).TFT, CLng(CurProp), ConvertFrom, DefaultUnit, False)
            For i = 1 To NumMethods
                If InfoMethod(CurProp).EqNum(i) <> 0 Then
                    InfoMethod(CurProp).MinT(i) = Convert(InfoMethod(CurProp).MinT(i), CLng(CurProp), ConvertFrom, DefaultUnit, False)
                    InfoMethod(CurProp).MaxT(i) = Convert(InfoMethod(CurProp).MaxT(i), CLng(CurProp), ConvertFrom, DefaultUnit, False)
                End If
            Next i
            InfoMethod(CurProp).TFTUnit = DefaultUnit
            If Trim(ConvertTo) = Trim(DefaultUnit) Then GoTo NextProp:
            InfoMethod(CurProp).TFT = Convert(InfoMethod(CurProp).TFT, CLng(CurProp), DefaultUnit, ConvertTo, False)
            For i = 1 To NumMethods
                If InfoMethod(CurProp).EqNum(i) <> 0 Then
                    InfoMethod(CurProp).MinT(i) = Convert(InfoMethod(CurProp).MinT(i), CLng(CurProp), DefaultUnit, ConvertTo, False)
                    InfoMethod(CurProp).MaxT(i) = Convert(InfoMethod(CurProp).MaxT(i), CLng(CurProp), DefaultUnit, ConvertTo, False)
                End If
            Next i
            InfoMethod(CurProp).TFTUnit = ConvertTo
NextProp:
            CurProp = TempProp
        Next J
    Else
        For J = 0 To NumProperties
            TempProp = CurProp
            CurProp = J
            InfoMethod(CurProp).TFT = Convert(InfoMethod(CurProp).TFT, CLng(CurProp), ConvertFrom, ConvertTo, False)
            For i = 1 To NumMethods
                If InfoMethod(CurProp).EqNum(i) <> 0 Then
                    InfoMethod(CurProp).MinT(i) = Convert(InfoMethod(CurProp).MinT(i), CLng(CurProp), ConvertFrom, ConvertTo, False)
                    InfoMethod(CurProp).MaxT(i) = Convert(InfoMethod(CurProp).MaxT(i), CLng(CurProp), ConvertFrom, ConvertTo, False)
                End If
            Next i
            InfoMethod(CurProp).TFTUnit = ConvertTo
            CurProp = TempProp
        Next J
    End If
                
End Sub
Sub ConvertTFTUnits2(ConvertFrom As String, ConvertTo As String)
    
    Dim i As Integer
    Dim J As Integer
    Dim TempProp As Integer
    Dim DefaultUnit As String
    
    If Trim(ConvertFrom) = Trim(ConvertTo) Then Exit Sub
                
    DefaultUnit = Get_DefaultUnit(-1)
    
    If Trim(ConvertFrom) <> Trim(DefaultUnit) Then
        For J = 0 To NumProperties
            TempProp = CurProp
            CurProp = J
            For i = 1 To NumMethods
                If InfoMethod(CurProp).EqNum(i) <> 0 Then
                    InfoMethod(CurProp).MinT(i) = Convert(InfoMethod(CurProp).MinT(i), CLng(CurProp), ConvertFrom, DefaultUnit, False)
                    InfoMethod(CurProp).MaxT(i) = Convert(InfoMethod(CurProp).MaxT(i), CLng(CurProp), ConvertFrom, DefaultUnit, False)
                End If
            Next i
            InfoMethod(CurProp).TFTUnit = DefaultUnit
            If Trim(ConvertTo) = Trim(DefaultUnit) Then
                Exit Sub
            End If
            For i = 1 To NumMethods
                If InfoMethod(CurProp).EqNum(i) <> 0 Then
                    InfoMethod(CurProp).MinT(i) = Convert(InfoMethod(CurProp).MinT(i), CLng(CurProp), DefaultUnit, ConvertTo, False)
                    InfoMethod(CurProp).MaxT(i) = Convert(InfoMethod(CurProp).MaxT(i), CLng(CurProp), DefaultUnit, ConvertTo, False)
                End If
            Next i
            InfoMethod(CurProp).TFTUnit = ConvertTo
            CurProp = TempProp
        Next J
    Else
        For J = 0 To NumProperties
            TempProp = CurProp
            CurProp = J
            For i = 1 To NumMethods
                If InfoMethod(CurProp).EqNum(i) <> 0 Then
                    InfoMethod(CurProp).MinT(i) = Convert(InfoMethod(CurProp).MinT(i), CLng(CurProp), ConvertFrom, ConvertTo, False)
                    InfoMethod(CurProp).MaxT(i) = Convert(InfoMethod(CurProp).MaxT(i), CLng(CurProp), ConvertFrom, ConvertTo, False)
                End If
            Next i
            InfoMethod(CurProp).TFTUnit = ConvertTo
            CurProp = TempProp
        Next J
    End If
                
End Sub


Sub LoadUnitsForm()

    Dim i As Integer
       
    FRMUnits!CMBUnits.Clear
    If CurProp >= 0 And CurProp <= 54 Then
        Call fill_units_form(CInt(CurProp))
        Exit Sub
    End If
    FRMUnits!CMBUnits.AddItem Get_DefaultUnit(CurProp)

    i = 1
    Do While Unit1(i) <> "End"
        If Trim(Unit1(i)) = Trim(Get_DefaultUnit(CurProp)) Then
            FRMUnits!CMBUnits.AddItem Unit2(i)
        
        ElseIf Trim(Unit2(i)) = Trim(Get_DefaultUnit(CurProp)) Then
            FRMUnits!CMBUnits.AddItem Unit1(i)
        End If
        i = i + 1
    Loop
    
    If CurProp = OptTemp Then
        FRMUnits!CMBUnits.Text = Cur_Disp.OpTUnit
    ElseIf CurProp = OptPress Then
        FRMUnits!CMBUnits.Text = Cur_Disp.OpPUnit
    Else
        FRMUnits!CMBUnits.Text = DispMethod(CurProp).Unit
    End If

End Sub

Sub ConvertUnits(ConvertFrom As String, ConvertTo As String)
    
    Dim i As Integer
    Dim tempLD As Double
    Dim DefaultUnit As String
    Dim temp_answer As Double
    
    If Trim(ConvertFrom) = Trim(ConvertTo) Then Exit Sub
    
    DefaultUnit = Get_DefaultUnit(CurProp)
    
    If CurProp = OptTemp Then
        If Trim(ConvertFrom) <> Trim(DefaultUnit) Then
            Cur_Disp.OpT = Convert(Cur_Disp.OpT, OptTemp, ConvertFrom, DefaultUnit, True)
            Cur_Disp.OpTUnit = DefaultUnit
            If Trim(ConvertTo) = Trim(DefaultUnit) Then
                Exit Sub
            End If
            Cur_Disp.OpT = Convert(Cur_Disp.OpT, OptTemp, ConvertFrom, ConvertTo, True)
            Cur_Disp.OpTUnit = ConvertTo
        Else
            Cur_Disp.OpT = Convert(Cur_Disp.OpT, OptTemp, ConvertFrom, ConvertTo, True)
            Cur_Disp.OpTUnit = ConvertTo
        End If
    ElseIf CurProp = OptPress Then
        If Trim(ConvertFrom) <> Trim(DefaultUnit) Then
            Cur_Disp.OpP = Convert(Cur_Disp.OpP, OptPress, ConvertFrom, DefaultUnit, True) 'FIX no code for OpP
            Cur_Disp.OpPUnit = DefaultUnit
            If Trim(ConvertTo) = Trim(DefaultUnit) Then
                Exit Sub
            End If
            Cur_Disp.OpP = Convert(Cur_Disp.OpP, OptPress, ConvertFrom, ConvertTo, True) 'FIX no code for OpP
            Cur_Disp.OpPUnit = ConvertTo
        Else
            Cur_Disp.OpP = Convert(Cur_Disp.OpP, OptPress, ConvertFrom, ConvertTo, True) 'FIX no code for OpP
            Cur_Disp.OpPUnit = ConvertTo
        End If
    Else    ' most of "else" commented out 6/30/97 BGH - this
            ' extensive code is not needed for the conversion.
            For i = 1 To NumMethods
                    If DispMethod(CurProp).Enabled(i) = True Then
                    
                        temp_answer = Convert(DispMethod(CurProp).value(i), CLng(CurProp), ConvertFrom, ConvertTo, True)
                        If temp_answer <> ERROR_FLAG Then
                            DispMethod(CurProp).value(i) = temp_answer
                            
                        Else
                            Exit Sub
                        End If
                    End If
            Next i
            
            DispMethod(CurProp).Unit = ConvertTo
    End If
                
End Sub






