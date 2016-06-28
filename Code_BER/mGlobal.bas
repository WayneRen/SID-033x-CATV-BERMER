Attribute VB_Name = "mGlobal"

Global gSID As String

Global gstrCaption As String

Global Const gconMAX_BAND = 6

'===== SQL Table =====
Global Const gconSQLTable = "TXBERTD"


'***** Path Def
Public Type typPath
  sApp As String
  sCurDir As String
  sConfig As String
  sTestData As String
  '***** L.Ly 05/19/2014
  sComputerName As String
End Type
Global gPATH As typPath


'***** Spec (Arrays)
Public Type typSpec
  iIndex() As Integer
  sModel() As String
  sPart() As String
  sATP() As String
  sPS() As String
  sDescription() As String
  sInstructions() As String
  '***
  sBER_MHz() As String
  sBER_Min() As String
  sBER_Max() As String
  '**
  sMER_Min() As String
  sMER_Max() As String
  '**
  dGate_Sec() As Double
  '**
  sWL_Table() As String
  '**
  sOffset() As String
End Type
Global gSPEC As typSpec


'***** Test Type (Arrays)
Public Type typTEST_TYPE
  sTest_Type() As String
End Type
Global gTEST_TYPE As typTEST_TYPE


Public Type typGPIB
  iBER As Integer     'LW8200
End Type
Public gGPIB As typGPIB

Public gblnEQInit As Boolean

'===== EQUIPMENT Control Class/Vars/Type
'Equipment class
Public Type typEQ
  iError As Integer
  '***** GPIB Classes
  clsBER As New clsEFA1500
End Type
Public gEQ As typEQ


'===== Station Info
Public Type typStation
  sStationID As String
  sTestLoc As String
  sDBaseType As String
End Type
Global gSTA As typStation


Global giIndex As Integer

'Global Const Pass_Color = &H80000009
Global Const Pass_Color = &HC0FFC0
Global Const Fail_Color = &HC0C0FF
Global Const White_Color = &H80000004


Global giTimerCnt As Integer

Global giCurrBand As Integer

Public Type typDBase
  '*** SQL
  SQLServer As String
  SQLDatabase As String
  SQLPassword As String
  SQLUser As String
End Type
Global gDB As typDBase

Public gsKeyDBase() As String

Public gsKeyStation() As String

Public gbPass As Boolean

Public gbStop As Boolean

Public gOffset As String



'=========DLL Declarations

'***** Added by L.Ly 05/19/2014
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long



