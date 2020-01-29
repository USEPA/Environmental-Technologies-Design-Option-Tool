Attribute VB_Name = "Structs"
Option Explicit


Type FontInfoType
  FontBold As Boolean
  FontItalic As Boolean
  FontName As String
  FontSize As Double
  FontStrikeThru As Boolean
  FontUnderline As Boolean
End Type
Type ColorInfoType
  Color As Long
End Type
Type PositionInfoType
  Left As Double
  Top As Double
  Width As Double
  Height As Double
End Type
'---------------------


Type IconType
  Name As String
  LongName As String
  DescriptionText As String
  fn_IconImage As String
  fn_ApplicationLink As String
  fn_ApplicationLink_Dir As String
End Type

Type GroupType
  Name As String
  Icons() As IconType
  Icons_Count As Integer
  GroupBackgroundColor As ColorInfoType
  GroupForegroundColor As ColorInfoType
  GroupTitleFont As FontInfoType
  GroupIconFont As FontInfoType
  Pos As PositionInfoType
End Type

Type TabType
  Name As String
  Groups() As GroupType
  Groups_Count As Integer
  fn_BackgroundImage As String
  TabBackgroundColor As ColorInfoType
End Type

'SEE ALSO: MainDatafile_Input, MainDatafile_Output
Type ProjectType
  dirty As Boolean
  Tabs() As TabType
  Tabs_Count As Integer
  Pos As PositionInfoType
  lvGroups_View As Integer                        '(I)
  lvGroups_Arrange As Integer                     '(I)
  ShowDescriptionText As Boolean                  '(I)
  MinimizeOnApplicationExecution As Boolean       '(I)
  StartupMode As Integer                          '(I)
  DisplayUninstalledApplications As Boolean       '(I)
  'NOTES:
  '  (I) = STORED IN .INI FILE INSTEAD OF CPAS.DAT FILE.
End Type

Global NowProj As ProjectType

  
  
