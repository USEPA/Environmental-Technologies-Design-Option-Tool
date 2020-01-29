Attribute VB_Name = "C_PropMod"
Option Explicit

'Capabilities of frmContaminantPropertyEdit:
'- Start editing all components at component #X.
'- Add one component.
'- Both of these were cancelled to perform a StEPP-Import.

Type rec_frmContaminantPropertyEdit
  'Inputs.
  ModelName As String         'Used only for caption
  ModelType As Integer        'Used to determine which properties are edited/displayed
  DoEditNumber As Integer     'Start editing on this component
  DoAdd As Integer            'Add one component
  OldNumCompo As Integer      'Old number of components

  'Input and Output.
  Contaminants(MAXCHEMICAL) As ContaminantPropertyType
  
  'Outputs.
  StEPPImportedNum As Integer 'Imported X components
  CancelledEdit As Integer    'User cancelled edit
  CancelledAdd As Integer     'User cancelled addition
  NewNumCompo As Integer      'New number of components
End Type

Type rec_Units_frmContaminantPropertyEdit
  'Units on frmContaminantPropertyEdit.
  UnitsProp(0 To 5) As String
  UnitsConc(0 To 1) As String
End Type

Global Data_frmContaminantPropertyEdit As rec_frmContaminantPropertyEdit
Global Units_frmContaminantPropertyEdit As rec_Units_frmContaminantPropertyEdit

Global Const MODELTYPE_PACKEDTOWER = 1
Global Const MODELTYPE_BUBBLE = 2
Global Const MODELTYPE_SURFACE = 3

