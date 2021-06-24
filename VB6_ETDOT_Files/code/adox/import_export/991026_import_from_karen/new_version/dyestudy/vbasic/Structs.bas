Attribute VB_Name = "Structs"
Option Explicit

Global frmPrint_DO_INPUTS As Boolean
Global frmPrint_DO_OUTPUTS As Boolean
Global frmPrint_DO_PLOTS As Boolean

Global Const USE_FONTNAME = "arial"
Global Const USE_FONTSIZE = 8
Global Const USE_FORMAT_CURRENCYSTANDARD = "$#,##0_);[Red]($#,##0)"
Global Const USE_FORMAT_CURRENCYDIGITSPAST2 = "$#,##0.00_);[Red]($#,##0.00)"

Global Const Latest_DataVersion_Major = 1
Global Const Latest_DataVersion_Minor = 0
Global Current_Filename As String
Global IsCalculated As Boolean
Global Predicted_Available As Boolean

Type DyeStudy_Type
  time As String
  concentration As String
End Type
  
Type Predicted_Type
  Predicted_Theta As Double
  Predicted_E As Double
End Type

Type Experimental_Type
  Experimental_Theta As Double
  Experimental_E As Double
End Type

Type Project_Type
  Filename As String
  dirty As Boolean          'has any data changed?
  'DYE STUDY PARAMETERS.
  dyestudy_output As String
  dyestudy_count As Integer
  DyeStudy() As DyeStudy_Type
  dyestudy_calcdate As String
  Predicted_Available As Boolean
  Predicted_count As Integer
  Predicted() As Predicted_Type
  Experimental_count As Integer
  Experimental() As Experimental_Type
End Type


Global nowproj As Project_Type
