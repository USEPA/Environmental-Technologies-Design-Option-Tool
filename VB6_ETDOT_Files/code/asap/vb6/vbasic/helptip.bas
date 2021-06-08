Attribute VB_Name = "HelpTipMod"
Option Explicit
'Declarations for HelpTips

Type PointType
  x As Integer
  Y As Integer
End Type

Declare Function GetActiveWindow Lib "User" () As Integer
Declare Sub GetCursorPos Lib "User" (PointStructure As PointType)
Declare Function ShowWindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function WindowFromPoint Lib "User" (ByVal PointStructY As Integer, ByVal PointStructX As Integer) As Integer
Global Const SW_SHOWNOACTIVE = 4


'Declarations for Displaying Menu Prompts

'
' Message sent by windows when a menu is selected
'
Global Const WM_MENUSELECT = &H11F
'
' Windows API Functions
'
Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetMenuItemID Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
'
' Used to locate prompt string for a menu
'
Type MenuPromptMap
    menuID  As Integer
    prompt  As String
End Type
'
' Room for 100 menu prompts
'
Global menuPrompts(100) As MenuPromptMap
'
' Contains index of last menu prompt string added to array
'
Global iMenuPrompts     As Integer

Sub ShowHelpTip(TipText$)
' Dim PointStruct As PointType
' Dim TopOffset As Integer
' Dim LeftOffset As Integer
' Dim r%
'
' If Len(TipText$) <> 0 Then
'   HelpTipForm.Hide
'   HelpTipForm.HelpTipLabel.Caption = TipText$
'   Call GetCursorPos(PointStruct)
'   TopOffset = 18
'   LeftOffset = -2
'
'   HelpTipForm.Width = HelpTipForm.HelpTipLabel.Width + 4 * Screen.TwipsPerPixelX
'   HelpTipForm.Height = HelpTipForm.HelpTipLabel.Height + 2 * Screen.TwipsPerPixelY
'
'
'   HelpTipForm.Top = (PointStruct.Y + TopOffset) * Screen.TwipsPerPixelY
'   HelpTipForm.Left = (PointStruct.x + LeftOffset) * Screen.TwipsPerPixelX
'
''   HelpTipForm.Width = HelpTipForm.HelpTipLabel.Width + 4 * Screen.TwipsPerPixelX
''   HelpTipForm.Height = HelpTipForm.HelpTipLabel.Height + 2 * Screen.TwipsPerPixelY
'
'   HelpTipForm.ZOrder
'   r% = ShowWindow(HelpTipForm.hWnd, SW_SHOWNOACTIVE)
' Else
'   HelpTipForm.Hide
' End If
End Sub

