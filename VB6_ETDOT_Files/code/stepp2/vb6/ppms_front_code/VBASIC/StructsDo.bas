Attribute VB_Name = "StructsDo"
Option Explicit



Const StructsDo_declarations_end = True


Sub Project_SetDefaults(Prj As Project_Type)
  Prj.length = 1#
  Prj.Diameter = 1#
  Prj.Mass = 1#
  Prj.FlowRate = 1#
End Sub


