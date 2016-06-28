Attribute VB_Name = "mGeneral"


Public Function CheckLimits(ByVal sSpecPair As String, _
                    ByVal sMeas As String, _
                    ByRef sResult As String) As Boolean
'*************************
'* sSpecPair: - "Low/High", ">= Min", "<=Max"
'*
'* Return: - PASS = True
'*         - FAIL = False
'*************************
  
  Dim lRetPos As Long
  Dim dLow As Double
  Dim dHi As Double
    
  CheckLimits = True
    
  '***** Low/High pair
  lRetPos = InStr(1, sSpecPair, "/")
  If lRetPos > 0 Then
    dLow = Val(Left(sSpecPair, lRetPos - 1))
    dHi = Val(Right(sSpecPair, Len(sSpecPair) - lRetPos))
    sResult = sMeas
    If Val(sResult) >= dLow And _
        Val(sResult) <= dHi Then
      CheckLimits = True
    Else
      CheckLimits = False
      sResult = sResult & "*"
    End If
    Exit Function
  End If
  
  '***** (<=) High limit only
  lRetPos = InStr(1, sSpecPair, "<=")
  If lRetPos > 0 Then
    dHi = Val(Right(sSpecPair, Len(sSpecPair) - lRetPos - 1))
    sResult = sMeas
    If Val(sResult) <= dHi Then
      CheckLimits = True
    Else
      CheckLimits = False
      sResult = sResult & "*"
    End If
    Exit Function
  End If
    
  '***** (<=) Low limit only
  lRetPos = InStr(1, sSpecPair, ">=")
  If lRetPos > 0 Then
    dLow = Val(Right(sSpecPair, Len(sSpecPair) - lRetPos - 1))
    sResult = sMeas
    If Val(sResult) >= dLow Then
      CheckLimits = True
    Else
      CheckLimits = False
      sResult = sResult & "*"
    End If
    Exit Function
  End If

End Function

Public Sub FileToKeyArray(ByVal sFile As String, sRetString() As String)

  Dim iFileNum As Integer
  Dim iLineCnt As Integer
  
  
  On Error GoTo ErrorHandler
  
  'ReDim sRetString(1000)
  ReDim sRetString(3000)
  
  iFileNum = FreeFile
  
  Open sFile For Input As #iFileNum
  
  iLineCnt = 0
  Do Until EOF(iFileNum)
      Line Input #iFileNum, sRetString(iLineCnt)
      sRetString(iLineCnt) = Trim(sRetString(iLineCnt))
      iLineCnt = iLineCnt + 1
  Loop
  
  Close #iFileNum
  
  ReDim Preserve sRetString(iLineCnt)
  

  Exit Sub

ErrorHandler:
  Select Case Err
  Case 53 'file not found
      MsgBox sFile & " not found."
  Case Else
      MsgBox "Error:  " & Error(Err)
  End Select

End Sub


Public Sub LoadTestLoc()

  Dim sBuf As String
  
  'Dim sKey() As String
  

  '------------------ Station ID etc...

  sBuf = gPATH.sConfig & "\STNID.TXT"
  
  Call mGeneral.FileToKeyArray(sBuf, gsKeyStation)
  
  gSTA.sStationID = mGeneral.ReadKeyArray(gsKeyStation, "Station", "Name")
  
  gSTA.sDBaseType = mGeneral.ReadKeyArray(gsKeyStation, "Station", "DBase_Type")

  gSTA.sTestLoc = mGeneral.ReadKeyArray(gsKeyStation, "Station", "Test_Loc")


End Sub

Public Function ReadKeyArray(sStringArray() As String, _
                        ByRef sSection As String, _
                        ByRef sKey As String) As String
                        
  
  
  Dim iCnt As Integer
  Dim iSection As Integer
  Dim iKey As Integer
  
  Dim iIndex As Integer
  Dim sBuf As String
  
  Dim sKeyName As String
  
  
  ReadKeyArray = ""
  
  '***** Section search
  iSection = -1
  For iCnt = 0 To UBound(sStringArray)
    If Left(sStringArray(iCnt), 1) <> "'" Then
      If InStr(1, sStringArray(iCnt), "[" & sSection & "]", vbTextCompare) > 0 Then
        iSection = iCnt
        Exit For
      End If
    End If
  Next iCnt
      
  '***** Key search
  If iSection >= 0 Then
    iKey = -1
    For iCnt = (iSection + 1) To UBound(sStringArray)
      '*** Is 1st Char is ' --- Comment out char
      '***   or [ --- Next Section?
      Select Case Left(sStringArray(iCnt), 1)
      Case "'"
        'No action
      Case "["
        Exit For
      Case Else
        '*** 3 diff search: '=', 'Tab', and 'space'
        iIndex = InStr(1, sStringArray(iCnt), "=", vbTextCompare)
        If iIndex > 1 Then
          ' '=' found
        Else
          iIndex = InStr(1, sStringArray(iCnt), vbTab, vbTextCompare)
          If iIndex > 1 Then
            ' 'Tab' found
          Else
            iIndex = InStr(1, sStringArray(iCnt), " ", vbTextCompare)
            If iIndex > 1 Then
              ' 'Space' found
            End If
          End If
        End If
        '***
        If iIndex > 1 Then
          sKeyName = Trim(Left(sStringArray(iCnt), iIndex - 1))
          If Trim(UCase(sKey)) = UCase(sKeyName) Then
            iKey = iCnt
            sBuf = sStringArray(iCnt)
            Exit For
          End If
        End If
      End Select
    Next iCnt
    '***
    If iKey >= 0 Then
      '*** 3 diff search: '=', 'Tab', and 'space'
      iIndex = InStr(1, sStringArray(iCnt), "=", vbTextCompare)
      If iIndex > 1 Then
        ' '=' found
      Else
        iIndex = InStr(1, sStringArray(iCnt), vbTab, vbTextCompare)
        If iIndex > 1 Then
          ' 'Tab' found
        Else
          iIndex = InStr(1, sStringArray(iCnt), " ", vbTextCompare)
          If iIndex > 1 Then
            ' 'Space' found
          End If
        End If
      End If
      '***
      ReadKeyArray = Trim(Mid(sBuf, iIndex + 1, Len(sBuf)))
    End If
  End If

End Function


Public Sub RevisionHistory()

  '********************************
  '* Revision History
  '********************************
  
  'gSID = "SID-03xx Rev X1"  ' @ X1 --- Based on LEO s/w concept
                            ' for slower PC.
  
  gSID = "SID-03xx Rev A1" 'Added offset according to ME's request, change the revision to A1
End Sub

