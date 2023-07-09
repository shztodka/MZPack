Attribute VB_Name = "GlobalSettingsAndVariables"
'***********************************************
'           MZPack - Excel Utility Pack
'        2012, Katarzyna Zbroinska
'             katarzyna@zbroinski.net
'              www.zbroinski.net
'
' This software is distributed under the term of
'     General Public License (GNU) version 3.
'***********************************************

Option Explicit

'--------
'Settings
'--------

Public Const MYAPPNAME As String = "ABACEUP"
Public Const REGISTRY_SETTINGS As String = "SETTINGS"
Public Const REGISTRY_SETTING_REFERENCE_STYLE As String = "ReferenceStyle"
Public Const VERSION As String = "2.00"


'---------
'Variables
'---------

Public blnScreenUpdatingStatus As Boolean
Public enmCalculationStatus As XlCalculation
Public blnEnableEventsStatus As Boolean
