Attribute VB_Name = "mSQL"
Public Sub InitSQLDBase()

  Dim sFileName As String
  
  Dim sBuf As String
  
  Dim hdlFile As Integer    'File handle
  
  Dim iCnt As Integer
  
  
  sFileName = gPATH.sConfig & "\SQLSetup.txt"

  hdlFile = FreeFile
  
  Open sFileName For Output As #hdlFile
  
  Print #hdlFile, "'*************************************"
  Print #hdlFile, "'*  Data Base (SQL) set-up File"
  Print #hdlFile, "'*"
  Print #hdlFile, "'*  [Section]"
  Print #hdlFile, "'*  key_1 = Value"
  Print #hdlFile, "'*  key_2 = Value"
  Print #hdlFile, "'*"
  Print #hdlFile, "'*  Use ' as 1st Char for comment"
  Print #hdlFile, "'*"
  Print #hdlFile, "'*    [Section] : [ORA], [SQL]"
  Print #hdlFile, "'*"
  Print #hdlFile, "'*    Key: Database, Password, Server, User..."
  Print #hdlFile, "'*"
  Print #hdlFile, "'*"
  Print #hdlFile, "'*************************************"
  Print #hdlFile, ""
  Print #hdlFile, ""
  Print #hdlFile, "[SQL]"
  Print #hdlFile, "'***** SQL Set up"
  Print #hdlFile, "Server = CAALHSDB01"
  Print #hdlFile, "'Server = HBLFCSDB01.HBLFC.EMCORE.CN"
  Print #hdlFile, "Database = OrtelTE"
  Print #hdlFile, "Password = ortel"
  Print #hdlFile, "User = netuser"
  Print #hdlFile, ""
  
  
  Close #hdlFile
  
End Sub



Public Sub LoadSpec()
  
  Dim rs As New ADODB.Recordset
  Dim sConn As String
  Dim sSQL As String
  
  Dim iMax As Integer
  
  Dim iCnt As Integer
  
  Dim iBand As Integer
  Dim iBand2 As Integer
  
  Dim sBuf As String
  Dim sBuf2 As String
  
  
  
  'Connection string
  sConn = "Provider=Microsoft.Jet.OLEDB.4.0"
  sConn = sConn & ";Data Source=" & gPATH.sConfig & "\Spec_BER.mdb"
  sConn = sConn & ";Persist Security Info=False"
  
  
  '------------- Load Specs
  
  'SQL string
  sSQL = ""
  sSQL = sSQL & "SELECT * FROM [Spec] "
  sSQL = sSQL & "ORDER BY [Index] ASC "
  'Open Record
  rs.Open sSQL, sConn, adOpenKeyset, adLockReadOnly

  If rs.EOF = True Then
    MsgBox "Wrong Model??? --- No tests!!!"
    End
  End If
  
  iMax = rs.RecordCount
  
  With gSPEC
    ReDim .iIndex(iMax)
    ReDim .sATP(iMax)
    ReDim .sDescription(iMax)
    ReDim .sInstructions(iMax)
    ReDim .sModel(iMax)
    ReDim .sPart(iMax)
    ReDim .sPS(iMax)
    '***
    ReDim .sBER_MHz(6, iMax)
    ReDim .sBER_Min(6, iMax)
    ReDim .sBER_Max(6, iMax)
    '**
    ReDim .sMER_Min(6, iMax)
    ReDim .sMER_Max(6, iMax)
    '**
    ReDim .sOffset(6, iMax)
    
    ReDim .dGate_Sec(iMax)
    
    ReDim .sWL_Table(iMax)
  
  End With
  
  
  For iCnt = 1 To iMax
    
    With gSPEC
      .iIndex(iCnt) = rs("Index")
      .sATP(iCnt) = rs("ATP")
      .sDescription(iCnt) = rs("Description")
      .sInstructions(iCnt) = rs("Instructions")
      .sModel(iCnt) = rs("Model")
      .sPart(iCnt) = rs("PartNum")
      .sPS(iCnt) = rs("PS")
      .dGate_Sec(iCnt) = rs("Gate_Sec")
      '*****
      .sWL_Table(iCnt) = rs("WL_Pwr_Table")
      
    End With
    
    For iBand = 1 To 6
      sBuf = "F" & iBand & "_MHz"
      If rs(sBuf) <> "" Then
        With gSPEC
          .sBER_MHz(iBand, iCnt) = rs(sBuf)
          sBuf = "F" & iBand & "_BER_Min"
          If rs(sBuf) <> "" Then .sBER_Min(iBand, iCnt) = rs(sBuf)
          sBuf = "F" & iBand & "_BER_Max"
          If rs(sBuf) <> "" Then .sBER_Max(iBand, iCnt) = rs(sBuf)
          '***
          sBuf = "F" & iBand & "_MER_Min"
          If rs(sBuf) <> "" Then .sMER_Min(iBand, iCnt) = rs(sBuf)
          sBuf = "F" & iBand & "_MER_Max"
          If rs(sBuf) <> "" Then .sMER_Max(iBand, iCnt) = rs(sBuf)
        End With
      End If
    Next iBand
    
    For iBand2 = 1 To 6
        sBuf2 = "F" & iBand2 & "_MER_Offset"
        If rs(sBuf2) <> "" Then
            gSPEC.sOffset(iBand2, iCnt) = rs(sBuf2)
        End If
    Next iBand2
       
      
    If iCnt < iMax Then
      rs.MoveNext
    End If
  
  Next iCnt
  
  rs.Close
  delay (0.5)
  
  
  '------------- Load Test Type
  
  'SQL string
  sSQL = ""
  sSQL = sSQL & "SELECT * FROM [Test_Type] "
  sSQL = sSQL & "ORDER BY [Test_Type] ASC "
  'Open Record
  rs.Open sSQL, sConn, adOpenKeyset, adLockReadOnly

  If rs.EOF = True Then
    MsgBox "Wrong Test Type??? --- No tests!!!"
    End
  End If
  
  iMax = rs.RecordCount
  
  With gTEST_TYPE
    ReDim .sTest_Type(iMax)
  End With
  
  For iCnt = 1 To iMax
    
    With gTEST_TYPE
      .sTest_Type(iCnt) = rs("Test_Type")
    End With
  
    If iCnt < iMax Then
      rs.MoveNext
    End If
  
  Next iCnt
  
  rs.Close
  
  
End Sub


Public Sub LoadWLPwr()
  
  Dim rs As New ADODB.Recordset
  Dim sConn As String
  Dim sSQL As String
  
  Dim iMax As Integer
  
  Dim iCnt As Integer
  
  Dim iBand As Integer
  
  Dim sBuf As String
  
  
  
  'Connection string
  sConn = "Provider=Microsoft.Jet.OLEDB.4.0"
  sConn = sConn & ";Data Source=" & gPATH.sConfig & "\Spec_BER.mdb"
  sConn = sConn & ";Persist Security Info=False"
  
  
  'SQL string
  sSQL = ""
  'sSQL = sSQL & "SELECT * FROM [Spec] "
  sSQL = sSQL & "SELECT * FROM [" & gSPEC.sWL_Table(giIndex) & "] "
  sSQL = sSQL & "WHERE [Test_Model] = '" & frmMain.cboTestModel.Text & "' "
  'Open Record
  rs.Open sSQL, sConn, adOpenKeyset, adLockReadOnly

  With frmMain
    If rs.EOF = True Then
      .lblSpec_WL.Caption = ""
      .lblSpec_Pwr.Caption = ""
    Else
      .lblSpec_WL.Caption = rs("WL_nm_Min") & "/" & rs("WL_nm_Max")
      .lblSpec_Pwr.Caption = rs("Pwr_dBm_Min") & "/" & rs("Pwr_dBm_Max")
    End If
     '***
    .txtWL.Text = ""
    .txtPwrDBM.Text = ""
  End With
  
  rs.Close
    
End Sub


Public Sub RecallData()
  
  Dim sConn As String
  Dim sSQL As String
  Dim adoRS As New ADODB.Recordset

  Dim sBuf As String
  
  Dim iBand As Integer
  
  Dim iCnt As Integer
  
  
  '***** Connection string for DBase
  If UCase(gSTA.sDBaseType) = "SQL" Then
    sConn = "Provider=SQLOLEDB.1;"
    sConn = sConn & "Persist Security Info=False;"
    sConn = sConn & "User ID=" & gDB.SQLUser & ";"
    sConn = sConn & "Password=" & gDB.SQLPassword & ";"
    sConn = sConn & "Initial Catalog=" & gDB.SQLDatabase & ";"
    sConn = sConn & "Data Source=" & gDB.SQLServer
  Else
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;"
    sConn = sConn & "Data Source=" & gPATH.sTestData & "\TxBER.mdb;"
    sConn = sConn & "Persist Security Info=False"
  End If
  
  '***** SQL string
  sSQL = ""
  sSQL = sSQL & "SELECT TOP 5 *  FROM [" & gconSQLTable & "] "
  sSQL = sSQL & "WHERE [SN] = '" & UCase(Trim(frmMain.txtSN)) & "'"
  sSQL = sSQL & "AND [TestType] = '" & frmMain.cboTestType.Text & "'"
  sSQL = sSQL & "ORDER BY [TestTime] DESC "
  
  '***** Open TABLE
  'adoRS.Open sSQL, sConn, adOpenKeyset, adLockOptimistic
  adoRS.Open sSQL, sConn, adOpenKeyset, adLockReadOnly
  
  '***** Record existed?
  If adoRS.EOF Then
    MsgBox "Record Not Found!!!"
    Exit Sub
  End If
  
  
  '----------- Main Model index???
  For iCnt = 1 To UBound(gSPEC.sModel)
    If UCase(gSPEC.sModel(iCnt)) = UCase(adoRS("Model")) Then
      giIndex = iCnt
      Exit For
    End If
  Next iCnt
  frmMain.cboModel.ListIndex = giIndex
  delay (1)
  frmMain.txtSN.Text = adoRS("SN")
  
  
  '----------- Test Model (P/N) index???
  frmMain.cboTestModel.ListIndex = 0
  For iCnt = 0 To (frmMain.cboTestModel.ListCount - 1)
    If UCase(adoRS("PN")) = UCase(frmMain.cboTestModel.List(iCnt)) Then
      frmMain.cboTestModel.ListIndex = iCnt
      Exit For
    End If
  Next iCnt
    
  
  For iBand = 1 To gconMAX_BAND
    If frmMain.lblBER_MHz(iBand).Visible = True Then
      frmMain.txtMeas_BER(iBand).Text = adoRS("BER" & iBand)
      frmMain.txtMeas_BER(iBand).BackColor = White_Color
      '***
      frmMain.txtMeas_MER(iBand).Text = adoRS("MER" & iBand)
      frmMain.txtMeas_MER(iBand).BackColor = White_Color
      '***
      frmMain.ProgressBar1(iBand).value = 0
    End If
  Next iBand
  
  '***** Display Record sets
  delay (2)
  Set frmMain.msgRecord.Recordset = adoRS
  delay (1)
  frmMain.msgRecord.Refresh
  delay (1)
  
  '***** Display Record Title
  'MainMenu.lblDBaseTitle = "GX2"
  frmMain.lblDBaseModel = adoRS("Model")
  frmMain.lblDBaseTestModel = adoRS("PN")
  frmMain.lblDBaseRecID = adoRS("RecordID")
  frmMain.lblDBasePassFail = adoRS("FailStatus")
  '***
  frmMain.lblDBaseWL = adoRS("OpticalWavelength1")
  frmMain.lblDBasePwrDBM = adoRS("OpticalPower")
  
  frmMain.Refresh
  
  '***** Close TABLE
  adoRS.Close
  delay (1)
  
   
End Sub


Public Sub SaveData()

  Dim sConn As String
  Dim sSQL As String
  Dim adoRS As New ADODB.Recordset

  Dim sBuf As String
  
  Dim iBand As Integer

  Dim lTestCnt As Integer


  '***** Connection string for DBase
  If UCase(gSTA.sDBaseType) = "SQL" Then
    sConn = "Provider=SQLOLEDB.1;"
    sConn = sConn & "Persist Security Info=False;"
    sConn = sConn & "User ID=" & gDB.SQLUser & ";"
    sConn = sConn & "Password=" & gDB.SQLPassword & ";"
    sConn = sConn & "Initial Catalog=" & gDB.SQLDatabase & ";"
    sConn = sConn & "Data Source=" & gDB.SQLServer
  Else
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;"
    sConn = sConn & "Data Source=" & gPATH.sTestData & "\TxBER.mdb;"
    sConn = sConn & "Persist Security Info=False"
  End If
  
  '***** SQL string
  sSQL = ""
  sSQL = sSQL & "SELECT TOP 5 *  FROM [" & gconSQLTable & "] "
  sSQL = sSQL & "WHERE [SN] = '" & UCase(Trim(frmMain.txtSN)) & "'"
  sSQL = sSQL & "ORDER BY [TestTime] DESC "
  
  '***** Open TABLE
  adoRS.Open sSQL, sConn, adOpenKeyset, adLockOptimistic
  
  '***** Record existed?
  If adoRS.EOF Then
    lTestCnt = 1
  Else
    'adoRS.MoveLast      'Comment out when in DESC order
    lTestCnt = adoRS("TestCount") + 1   'Increment test count
  End If
  
  adoRS.Close
  delay (2)
  
  '------------------- Re-open table without sorting
  '***** SQL string
  sSQL = ""
  sSQL = sSQL & "SELECT TOP 5 *  FROM [" & gconSQLTable & "] "
  sSQL = sSQL & "WHERE [SN] = '" & UCase(Trim(frmMain.txtSN)) & "'"
  
  '***** Open TABLE
  adoRS.Open sSQL, sConn, adOpenKeyset, adLockOptimistic
  
  '*****
  adoRS.AddNew
  
  
  '--------------- Save data -------------------
  
  '***** Info
  adoRS("SN") = UCase(frmMain.txtSN)
  adoRS("Model") = gSPEC.sModel(giIndex)
  'adoRS("PN") = gSPEC.sPart(giIndex)
  adoRS("PN") = frmMain.cboTestModel.Text
  adoRS("TestStation") = UCase(frmMain.txtStationID.Text)
  If frmMain.txtLotN.Text <> "" Then
    adoRS("LotN") = UCase(frmMain.txtLotN.Text)
  End If
  If frmMain.txtWO.Text <> "" Then
    adoRS("WO") = UCase(frmMain.txtWO.Text)
  End If
  adoRS("TestTime") = Now
  adoRS("Operator") = UCase(frmMain.txtOperator.Text)
  
  adoRS("TestType") = frmMain.cboTestType.Text
  
  adoRS("TestCount") = lTestCnt
  
  adoRS("PGM_NAME") = "SID-03xx"
  sBuf = Right(gSID, Len(gSID) - InStr(1, gSID, "Rev") + 1)
  adoRS("PGM_VER") = sBuf
  adoRS("FACILITY") = UCase(frmMain.txtTestLoc.Text)
  
  
  If gbPass = True Then
    adoRS("FailStatus") = UCase("PASS")
  Else
    adoRS("FailStatus") = UCase("FAIL")
  End If
  
  '***** Test Data
  For iBand = 1 To gconMAX_BAND
    If frmMain.lblBER_MHz(iBand).Visible = True Then
      adoRS("BER" & iBand) = Format(Val(frmMain.txtMeas_BER(iBand).Text), "0.00e+000")
      adoRS("MER" & iBand) = Format(Val(frmMain.txtMeas_MER(iBand).Text), "0.00e+000")
    End If
  Next iBand

  adoRS("OpticalWavelength1") = Val(frmMain.txtWL.Text)
  adoRS("OpticalPower") = Val(frmMain.txtPwrDBM.Text)
  
  
  adoRS.Update
  delay (2)
  '***** Close TABLE
  adoRS.Close
  delay (1)
  
  
  '------------------- Re-open table as read only
  '***** SQL string
  sSQL = ""
  sSQL = sSQL & "SELECT TOP 5 *  FROM [" & gconSQLTable & "] "
  sSQL = sSQL & "WHERE [SN] = '" & UCase(Trim(frmMain.txtSN)) & "'"
  sSQL = sSQL & "ORDER BY [TestTime] DESC "
 
  '***** Open TABLE
  adoRS.Open sSQL, sConn, adOpenKeyset, adLockReadOnly
  
  '***** Display Record sets
  delay (2)
  Set frmMain.msgRecord.Recordset = adoRS
  delay (1)
  frmMain.msgRecord.Refresh
  delay (1)
  
  '***** Display Record Title
  'MainMenu.lblDBaseTitle = "GX2"
  frmMain.lblDBaseModel = frmMain.cboModel.Text
  frmMain.lblDBaseTestModel = adoRS("PN")
  frmMain.lblDBaseRecID = adoRS("RecordID")
  frmMain.lblDBasePassFail = adoRS("FailStatus")
  '***
  frmMain.lblDBaseWL = adoRS("OpticalWavelength1")
  frmMain.lblDBasePwrDBM = adoRS("OpticalPower")
  
  frmMain.Refresh
  
  '***** Close TABLE
  adoRS.Close
  delay (1)
 
 
End Sub

