Attribute VB_Name = "Structs"
Option Explicit

Global Const Latest_DataVersion_Major = 1
Global Const Latest_DataVersion_Minor = 0
Global Current_Filename As String

Type Project_Type
  length As Double
  Diameter As Double
  Mass As Double
  FlowRate As Double
End Type
Global NowProj As Project_Type


