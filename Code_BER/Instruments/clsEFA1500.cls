VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEFA1500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'Private member variables
Private miDeviceDriverSet As Integer
Private miPrimaryAdd As Integer 'local copy
Private miSecondaryAdd As Integer 'local copy
Private miGpibBoard As Integer 'local copy
Private miEFA1500 As Integer
Private msErrorMess As String

Private mbDisplayMessON As Boolean



Public Property Let GpibBoard(ByVal vNewValue As Integer)
    ' write only property
    miGpibBoard = vNewValue
End Property

Public Function ReadQAMHistBER_Before_Ave(ByRef dBER As Double) As Integer
  
  Dim sCmd As String

  Dim sBuf As String
  
  ReadQAMHistBER_Before_Ave = 0
  
  sBuf = Space(20)
  
  sCmd = "READ:QAM:HISTORY:BER:BRS? AVG"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    DisplayMsg msErrorMess
    ReadQAMHistBER_Before_Ave = 1
    Exit Function
  End If
  
  'STOPend,
  Call Receive(miGpibBoard, miPrimaryAdd, sBuf, STOPend)  'read power
  'Call ibrd(miEFA1500, sBuf)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    ReadQAMHistBER_Before_Ave = 1
    Exit Function
  End If
  
  '***** Remove quote string????
  If Left(sBuf, 1) = """" Then
    sBuf = Mid(sBuf, 2, ibcnt)
  End If
  
  'sBER = Left(sBuf, ibcnt)
  
  'sBER = Format(sBuf, "0.00E-00")
  
  dBER = Val(sBuf)
    
End Function


Public Function ReadQAMHistBER_Before_Curr(ByRef dBER As Double) As Integer

  Dim sCmd As String

  Dim sBuf As String
  
  ReadQAMHistBER_Before_Curr = 0
  
  sBuf = Space(20)
  
  sCmd = "READ:QAM:HISTORY:BER:BRS? CUR"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    DisplayMsg msErrorMess
    ReadQAMHistBER_Before_Curr = 1
    Exit Function
  End If
  
  'STOPend,
  Call Receive(miGpibBoard, miPrimaryAdd, sBuf, STOPend)  'read power
  'Call ibrd(miEFA1500, sBuf)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    ReadQAMHistBER_Before_Curr = 1
    Exit Function
  End If
  
  '***** Remove quote string????
  If Left(sBuf, 1) = """" Then
    sBuf = Mid(sBuf, 2, ibcnt)
  End If
  
  'sBER = Left(sBuf, ibcnt)
  
  'sBER = Format(sBuf, "0.00E-00")
  
  dBER = Val(sBuf)

End Function


Public Function ReadQAMHistBER_Before_Max(ByRef dBER As Double) As Integer
 Dim sCmd As String

  Dim sBuf As String
  
  ReadQAMHistBER_Before_Max = 0
  
  sBuf = Space(20)
  
  sCmd = "READ:QAM:HISTORY:BER:BRS? MAX"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    DisplayMsg msErrorMess
    ReadQAMHistBER_Before_Max = 1
    Exit Function
  End If
  
  'STOPend,
  Call Receive(miGpibBoard, miPrimaryAdd, sBuf, STOPend)  'read power
  'Call ibrd(miEFA1500, sBuf)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    ReadQAMHistBER_Before_Max = 1
    Exit Function
  End If
  
  '***** Remove quote string????
  If Left(sBuf, 1) = """" Then
    sBuf = Mid(sBuf, 2, ibcnt)
  End If
  
  'sBER = Left(sBuf, ibcnt)
  
  'sBER = Format(sBuf, "0.00E-00")
  
  dBER = Val(sBuf)
End Function


Public Function ReadQAMHistBER_Before_Min(ByRef dBER As Double) As Integer

  Dim sCmd As String

  Dim sBuf As String
  
  ReadQAMHistBER_Before_Min = 0
  
  sBuf = Space(20)
  
  sCmd = "READ:QAM:HISTORY:BER:BRS? MIN"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    DisplayMsg msErrorMess
    ReadQAMHistBER_Before_Min = 1
    Exit Function
  End If
  
  'STOPend,
  Call Receive(miGpibBoard, miPrimaryAdd, sBuf, STOPend)  'read power
  'Call ibrd(miEFA1500, sBuf)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    ReadQAMHistBER_Before_Min = 1
    Exit Function
  End If
  
  '***** Remove quote string????
  If Left(sBuf, 1) = """" Then
    sBuf = Mid(sBuf, 2, ibcnt)
  End If
  
  'sBER = Left(sBuf, ibcnt)
  
  'sBER = Format(sBuf, "0.00E-00")
  
  dBER = Val(sBuf)

End Function


Public Function ReadQAMHistMER_DB_Curr(ByRef dMER As Double) As Integer
Dim sCmd As String

  Dim sBuf As String
  
  ReadQAMHistMER_DB_Curr = 0
  
  sBuf = Space(20)
  
  sCmd = "READ:QAM:HISTORY:MER:DB? CUR"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    DisplayMsg msErrorMess
    ReadQAMHistMER_DB_Curr = 1
    Exit Function
  End If
  
  'STOPend,
  Call Receive(miGpibBoard, miPrimaryAdd, sBuf, STOPend)  'read power
  'Call ibrd(miEFA1500, sBuf)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    ReadQAMHistMER_DB_Curr = 1
    DisplayMsg msErrorMess
    Exit Function
  End If
  
  '***** Remove quote string????
  If Left(sBuf, 1) = """" Then
    sBuf = Mid(sBuf, 2, ibcnt)
  End If
  
  'sBER = Left(sBuf, ibcnt)
  
  'sBER = Format(sBuf, "0.00E-00")
  
  dMER = Val(sBuf)
End Function

Public Function ReadQAMHistMER_DB_Max(ByRef dMER As Double) As Integer
  
  Dim sCmd As String

  Dim sBuf As String
  
  ReadQAMHistMER_DB_Max = 0
  
  sBuf = Space(20)
  
  sCmd = "READ:QAM:HISTORY:MER:DB? MAX"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    DisplayMsg msErrorMess
    ReadQAMHistMER_DB_Max = 1
    Exit Function
  End If
  
  'STOPend,
  Call Receive(miGpibBoard, miPrimaryAdd, sBuf, STOPend)  'read power
  'Call ibrd(miEFA1500, sBuf)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    ReadQAMHistMER_DB_Max = 1
    Exit Function
  End If
  
  '***** Remove quote string????
  If Left(sBuf, 1) = """" Then
    sBuf = Mid(sBuf, 2, ibcnt)
  End If
  
  'sBER = Left(sBuf, ibcnt)
  
  'sBER = Format(sBuf, "0.00E-00")
  
  dMER = Val(sBuf)
    
End Function



Public Function ReadQAMHistMER_DB_Ave(ByRef dMER As Double) As Integer
  
  Dim sCmd As String

  Dim sBuf As String
  
  ReadQAMHistMER_DB_Ave = 0
  
  sBuf = Space(20)
  
  sCmd = "READ:QAM:HISTORY:MER:DB? AVG"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    DisplayMsg msErrorMess
    ReadQAMHistMER_DB_Ave = 1
    Exit Function
  End If
  
  'STOPend,
  Call Receive(miGpibBoard, miPrimaryAdd, sBuf, STOPend)  'read power
  'Call ibrd(miEFA1500, sBuf)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    ReadQAMHistMER_DB_Ave = 1
    Exit Function
  End If
  
  '***** Remove quote string????
  If Left(sBuf, 1) = """" Then
    sBuf = Mid(sBuf, 2, ibcnt)
  End If
  
  'sBER = Left(sBuf, ibcnt)
  
  'sBER = Format(sBuf, "0.00E-00")
  
  dMER = Val(sBuf)
    
End Function

Public Function ReadQAMHistMER_DB_Min(ByRef dMER As Double) As Integer


  Dim sCmd As String
 

  Dim sBuf As String
  
  ReadQAMHistMER_DB_Min = 0
  
  sBuf = Space(20)
  
  sCmd = "READ:QAM:HISTORY:MER:DB? MIN"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    DisplayMsg msErrorMess
    ReadQAMHistMER_DB_Min = 1
    Exit Function
  End If
  
  'STOPend,
  Call Receive(miGpibBoard, miPrimaryAdd, sBuf, STOPend)  'read power
  'Call ibrd(miEFA1500, sBuf)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do ReadFreqStep"
    ReadQAMHistMER_DB_Min = 1
    Exit Function
  End If
  
  '***** Remove quote string????
  If Left(sBuf, 1) = """" Then
    sBuf = Mid(sBuf, 2, ibcnt)
  End If
  
  'sBER = Left(sBuf, ibcnt)
  
  'sBER = Format(sBuf, "0.00E-00")
  
  dMER = Val(sBuf)
    
End Function

Public Function SetQAMHistory_Display() As Integer
  
  Dim sCmd As String
  
  SetQAMHistory_Display = 0
  
  sCmd = "SENSE:QAM:HISTORY:DISPLAY NUMV"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do SetFreqFlag"
    DisplayMsg msErrorMess
    SetQAMHistory_Display = 1
    Exit Function
  End If
  
End Function

Public Function SetQAMHistory_Restart() As Integer
  
  Dim sCmd As String
  
  SetQAMHistory_Restart = 0
  
  sCmd = "SENSE:QAM:HISTORY:RESTART"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do SetFreqFlag"
    DisplayMsg msErrorMess
    SetQAMHistory_Restart = 1
    Exit Function
  End If

End Function


Private Sub Class_Initialize()
'******************************************************************************************
' Purpose:  Subprogram to initializae member variables
' Inputs:   None
' Outputs:  None
'******************************************************************************************

    miPrimaryAdd = 20
    miSecondaryAdd = 0
    miGpibBoard = 0
    miDeviceDriverSet = 0
    msErrorMess = ""
    
    'msModel = ""
    
    mbDisplayMessON = True
  
End Sub



Public Property Get PrimaryAdd() As Variant
  PrimaryAdd = miPrimaryAdd
End Property



Public Property Get ErrorMess() As String
    ErrorMess = msErrorMess
End Property



Public Property Let PrimaryAdd(ByVal vNewValue As Variant)
  miPrimaryAdd = vNewValue
End Property


Public Property Get SecondaryAdd() As Variant
  SecondaryAdd = miSecondaryAdd
End Property

Public Property Let SecondaryAdd(ByVal vNewValue As Variant)
  miSecondaryAdd = vNewValue
End Property


Public Function DisplayMsgOFF() As Integer

  mbDisplayMessON = False

End Function


Public Function DisplayMsgON() As Integer

  mbDisplayMessON = True

End Function


Public Function CloseDevice() As Integer
  Call ibonl(miEFA1500, 0)
  miDeviceDriverSet = 0
  CloseDevice = 0
End Function


Public Function OpenDevice() As Integer
  
  Dim addrlist%(2)
  
  addrlist%(1) = miPrimaryAdd
  addrlist%(2) = NOADDR
  
  OpenDevice = 0
  
  Call ibdev(miGpibBoard, miPrimaryAdd, miSecondaryAdd, T3s, 1, &HA, miEFA1500)         'Address
  If (ibsta And EERR) Then
    ' error detected
    OpenDevice = 1
    miDeviceDriverSet = 0
  Else
    'device opened ok
    OpenDevice = 0
    miDeviceDriverSet = 1
  End If
  'Call ibclr(miEFA1500)      'Set device in remote
  
  Call SendIFC(miGpibBoard) ' Reset GPIB by sending Interface clear
  If ibsta% And EERR Then
    DisplayMsg "Error in SendIFC"
    Exit Function
  End If
  
  Call DevClear(miGpibBoard, miPrimaryAdd)
  If ibsta% And EERR Then
    DisplayMsg "Error in DevClear"
    Exit Function
  End If
  
  
  Exit Function
  
  
  Call SendIFC(miGpibBoard) ' Reset GPIB by sending Interface clear
  If ibsta% And EERR Then
    DisplayMsg "Error in SendIFC"
    Exit Function
  End If
  
  Call DevClear(miGpibBoard, miPrimaryAdd)
  If ibsta% And EERR Then
    DisplayMsg "Error in DevClear"
    Exit Function
  End If
  
  
  Call EnableRemote(miGpibBoard, addrlist%())
  If ibsta% And EERR Then
    DisplayMsg "Error in EnableRemote"
    Exit Function
  End If
  
  
  
  
  Call SendIFC(miGpibBoard) ' Reset GPIB by sending Interface clear
  If ibsta% And EERR Then
    DisplayMsg "Error in SendIFC"
    Exit Function
  End If
  
  Call DevClear(miGpibBoard, miPrimaryAdd)
  If ibsta% And EERR Then
    DisplayMsg "Error in DevClear"
    Exit Function
  End If
  
  
  'Call ibdev(miGpibBoard, miPrimaryAdd, miSecondaryAdd, T3s, 1, 0, miEFA1500)  'Address
  Call ibdev(miGpibBoard, miPrimaryAdd, miSecondaryAdd, T3s, 1, &HA, miEFA1500)         'Address
  If (ibsta And EERR) Then
    ' error detected
    OpenDevice = 1
    miDeviceDriverSet = 0
  Else
    'device opened ok
    OpenDevice = 0
    miDeviceDriverSet = 1
  End If
  Call ibclr(miEFA1500)      'Set device in remote
  
End Function

Private Function DisplayMsg(ByVal sMessage As String) As Integer

  If mbDisplayMessON Then
    MsgBox sMessage, vbOKOnly, "EFA1500"
  End If
  
End Function



Public Function SetRFFrequencyMHz(ByVal dblMHz As Double) As Integer
  
  Dim sCmd As String

  SetRFFrequencyMHz = 0
  
  '*** Check input
  If dblMHz < 5 Then dblMHz = 5
  If dblMHz > 1000 Then dblMHz = 1000
  
  sCmd = "SENSE:FREQUENCY:RF " & dblMHz & " MHz"
  
  Call Send(miGpibBoard, miPrimaryAdd, sCmd, NLend)
  'Call ibwrt(miEFA1500, sCmd)
  If ibsta% And EERR Then
    msErrorMess = "488 ERROR - can't do SetFreqFlag"
    DisplayMsg msErrorMess
    SetRFFrequencyMHz = 1
    Exit Function
  End If
  
End Function
