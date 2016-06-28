Attribute VB_Name = "mEquip"
Public Sub GPIBInit()

  If gblnEQInit = False Then
    
    '***** BER
    If gGPIB.iBER < 1 Then
      'MsgBox "DVM GPIB Addr = -1 ---- Abort GPIBInit !!!"
    Else
      gEQ.clsBER.GpibBoard = 0
      gEQ.clsBER.PrimaryAdd = gGPIB.iBER
      gEQ.clsBER.SecondaryAdd = 0
      If gEQ.clsBER.OpenDevice Then
        MsgBox "BER OpenDevice failed  ---- Abort GPIBInit !!!"
        Exit Sub
      End If
    End If
    
  End If
  
  gblnEQInit = True
  
End Sub


Sub delay(delayTime!)
'//////////////////////////////////////////////////////////////////////////
'/
'/      Function:   Pauses for DelayTime! seconds
'/                  does not process events while pausing
'/
'/      Inputs:  This routine uses the following system globals:
'/                none
'/
'/      Outputs: None
'/      By:   Jerry Byrd
'/      Created:  2/14/96
'/      Last Rev: 2/14/96
'/
'/   (C) Copyright 1996  ORTEL Corp.  All Rights Reserved
'//////////////////////////////////////////////////////////////////////////

  Dim StopTime As Single
  
  StopTime = Timer + delayTime!
  Do Until Timer > StopTime
    DoEvents
    '***** Allow interruption?
    If gbStop = True Then
      Exit Do
    End If
  Loop

End Sub

