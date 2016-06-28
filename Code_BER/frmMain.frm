VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMain 
   Caption         =   "CATV BER Test   SID-03xx"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRecall 
      Caption         =   "Recall"
      Height          =   375
      Left            =   4680
      TabIndex        =   70
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   65
      Text            =   "Time"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   64
      Text            =   "Date"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtStationID 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   "txtStationID"
      ToolTipText     =   "Station Number"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtTestLoc 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "txtTestLoc"
      ToolTipText     =   "Test Location: LTK/LTC/EBB/LEO etc..."
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtDBaseType 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "txtDBaseType"
      ToolTipText     =   "DBase Type: SQL/Access"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2880
      TabIndex        =   47
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1560
      TabIndex        =   46
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdRUN 
      Caption         =   "Run"
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   43
      Text            =   "frmMain.frx":0000
      Top             =   2640
      Width           =   7815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5880
      Top             =   5400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   7815
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Index           =   1
         Left            =   720
         TabIndex        =   42
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.TextBox txtMeas_BER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   1
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "txtMeas_BER"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtMeas_MER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   1
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "txtMeas_MER"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtMeas_BER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "txtMeas_BER"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtMeas_MER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "txtMeas_MER"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtMeas_BER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   3
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "txtMeas_BER"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtMeas_MER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   3
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "txtMeas_MER"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtMeas_BER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   4
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "txtMeas_BER"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtMeas_MER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   4
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "txtMeas_MER"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtMeas_BER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   5
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "txtMeas_BER"
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtMeas_MER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   5
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "txtMeas_MER"
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtMeas_BER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   6
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "txtMeas_BER"
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtMeas_MER 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   6
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "txtMeas_MER"
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Index           =   2
         Left            =   1920
         TabIndex        =   48
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Index           =   3
         Left            =   3120
         TabIndex        =   49
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Index           =   4
         Left            =   4320
         TabIndex        =   50
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Index           =   5
         Left            =   5520
         TabIndex        =   51
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Index           =   6
         Left            =   6720
         TabIndex        =   52
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "BER"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "MHz"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblBER_MHz 
         BackColor       =   &H80000013&
         Caption         =   "lblBER_MHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSpec_BER 
         Caption         =   "lblSpec_BER"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   38
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "MER"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblSpec_MER 
         Caption         =   "lblSpec_MER"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblBER_MHz 
         BackColor       =   &H80000013&
         Caption         =   "lblBER_MHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSpec_BER 
         Caption         =   "lblSpec_BER"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSpec_MER 
         Caption         =   "lblSpec_MER"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblBER_MHz 
         BackColor       =   &H80000013&
         Caption         =   "lblBER_MHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSpec_BER 
         Caption         =   "lblSpec_BER"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSpec_MER 
         Caption         =   "lblSpec_MER"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   30
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblBER_MHz 
         BackColor       =   &H80000013&
         Caption         =   "lblBER_MHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSpec_BER 
         Caption         =   "lblSpec_BER"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   28
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSpec_MER 
         Caption         =   "lblSpec_MER"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   27
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblBER_MHz 
         BackColor       =   &H80000013&
         Caption         =   "lblBER_MHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5520
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSpec_BER 
         Caption         =   "lblSpec_BER"
         Height          =   255
         Index           =   5
         Left            =   5520
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSpec_MER 
         Caption         =   "lblSpec_MER"
         Height          =   255
         Index           =   5
         Left            =   5520
         TabIndex        =   24
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblBER_MHz 
         BackColor       =   &H80000013&
         Caption         =   "lblBER_MHz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSpec_BER 
         Caption         =   "lblSpec_BER"
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblSpec_MER 
         Caption         =   "lblSpec_MER"
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Meas"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Meas"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   375
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Info"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtPwrDBM 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   59
         ToolTipText     =   "Enter Opt Power in dBm"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtWL 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   58
         ToolTipText     =   "Enter WL in nm."
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox cboTestType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   73
         ToolTipText     =   "Select Correct Test Type"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboTestModel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   72
         ToolTipText     =   "Select Correct Test Model"
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtWO 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   57
         ToolTipText     =   "Enter Work Order Number?"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtLotN 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   56
         ToolTipText     =   "Enter Lot Number?"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtOperator 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   55
         ToolTipText     =   "Enter Operator Name or Initials."
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtSN 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   54
         ToolTipText     =   "Enter S/N here"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtGate_Sec 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "txtGate_Sec"
         ToolTipText     =   "Gate Time (Sec)"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtPartNum 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "txtPartNum"
         ToolTipText     =   "Product P/N"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtInstructions 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "txtInstructions"
         ToolTipText     =   "Product Instructions"
         Top             =   1080
         Width           =   9015
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "txtDescription"
         ToolTipText     =   "Product Description"
         Top             =   720
         Width           =   5055
      End
      Begin VB.ComboBox cboModel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Part Number"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblSpec_Pwr 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSpec_Pwr"
         Height          =   255
         Left            =   7200
         TabIndex        =   77
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblSpec_WL 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSpec_WL"
         Height          =   255
         Left            =   7200
         TabIndex        =   76
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Pwr dBm:"
         Height          =   255
         Left            =   5160
         TabIndex        =   75
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "WL nm:"
         Height          =   255
         Left            =   5160
         TabIndex        =   74
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "WO:"
         Height          =   255
         Left            =   2400
         TabIndex        =   67
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "LotN:"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Operator:"
         Height          =   255
         Left            =   2400
         TabIndex        =   61
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "S/N:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1560
         Width           =   615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msgRecord 
      Height          =   1335
      Left            =   120
      TabIndex        =   53
      ToolTipText     =   "Double click on the cell to recall test data."
      Top             =   6000
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2355
      _Version        =   393216
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblDBasePwrDBM 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDBasePwrDBM"
      Height          =   255
      Left            =   8880
      TabIndex        =   80
      ToolTipText     =   "DBase Pwr dBm"
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblDBaseWL 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDBaseWL"
      Height          =   255
      Left            =   7680
      TabIndex        =   79
      ToolTipText     =   "DBase WL"
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblDBaseTestModel 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDBaseModel"
      Height          =   255
      Left            =   6840
      TabIndex        =   78
      ToolTipText     =   "Test Model PN (from dBase record)"
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Label lblDBasePassFail 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDBasePassFail"
      Height          =   255
      Left            =   9000
      TabIndex        =   71
      ToolTipText     =   "DBase Pass/Fail"
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lblDBaseRecID 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDBaseRecID"
      Height          =   255
      Left            =   7680
      TabIndex        =   69
      ToolTipText     =   "Record ID"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblDBaseModel 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDBaseModel"
      Height          =   255
      Left            =   8040
      TabIndex        =   68
      ToolTipText     =   "Model (from dBase record)"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Menu mnuGPIB 
      Caption         =   "GPIB"
      Begin VB.Menu mnuGPIBErrorDisp 
         Caption         =   "GPIB Error Display ON"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboModel_Click()
  
  Dim iCnt As Integer
  Dim iBand As Integer
  
  Dim sLo As String
  Dim sHi As String
  
  Dim rs As New ADODB.Recordset
  Dim sConn As String
  Dim sSQL As String
  
  Dim iMax As Integer
  
  
  giIndex = Me.cboModel.ListIndex
  
  If giIndex = 0 Then
    With Me
      .txtPartNum.Text = ""
      .txtDescription.Text = ""
      .txtInstructions.Text = ""
      .txtGate_Sec.Text = ""
      '.txtOperator.Text = ""
      .txtSN.Text = ""
      '*****
      .cboTestModel.Clear
      .lblSpec_Pwr.Caption = ""
      .txtPwrDBM.Text = ""
      .lblSpec_WL.Caption = ""
      .txtWL.Text = ""
      '*****
    End With
    With Me
      For iBand = 1 To gconMAX_BAND
        .lblBER_MHz(iBand).Visible = False
        .lblSpec_BER(iBand).Visible = False
        .lblSpec_MER(iBand).Visible = False
        '***
        .txtMeas_BER(iBand).Visible = False
        .txtMeas_MER(iBand).Visible = False
        '***
        .ProgressBar1(iBand).Visible = False
      Next iBand
    End With
    Exit Sub
  End If
  
  With Me
    .txtPartNum.Text = gSPEC.sPart(giIndex)
    .txtDescription.Text = gSPEC.sDescription(giIndex)
    .txtInstructions.Text = gSPEC.sInstructions(giIndex)
    .txtGate_Sec.Text = gSPEC.dGate_Sec(giIndex)
    .txtSN.Text = ""
  End With
  
  '-------- Spec Assignment
  With Me
    For iBand = 1 To gconMAX_BAND
      '*****
      .lblBER_MHz(iBand).Visible = True
      .lblSpec_BER(iBand).Visible = True
      .lblSpec_MER(iBand).Visible = True
      .ProgressBar1(iBand).Visible = True
      .ProgressBar1(iBand).value = 0
      '***
      .txtMeas_BER(iBand).Visible = True
      .txtMeas_BER(iBand).Text = ""
      .txtMeas_BER(iBand).BackColor = Pass_Color
      .txtMeas_MER(iBand).Visible = True
      .txtMeas_MER(iBand).Text = ""
      .txtMeas_MER(iBand).BackColor = Pass_Color
      '*****
      If gSPEC.sBER_MHz(iBand, giIndex) <> "" Then
        .lblBER_MHz(iBand).Caption = gSPEC.sBER_MHz(iBand, giIndex)
        '***** BER
        sLo = gSPEC.sBER_Min(iBand, giIndex)
        sHi = gSPEC.sBER_Max(iBand, giIndex)
        If sLo <> "" And sHi <> "" Then
          .lblSpec_BER(iBand).Caption = sLo & "/" & sHi
        ElseIf sLo <> "" And sHi = "" Then
          .lblSpec_BER(iBand).Caption = ">= " & sLo
        ElseIf sLo = "" And sHi <> "" Then
          .lblSpec_BER(iBand).Caption = "<= " & sHi
        Else
          'No Spec ????
          .lblSpec_BER(iBand).Caption = "No spec?"
        End If
        '***** MER
        sLo = gSPEC.sMER_Min(iBand, giIndex)
        sHi = gSPEC.sMER_Max(iBand, giIndex)
        If sLo <> "" And sHi <> "" Then
          .lblSpec_MER(iBand).Caption = sLo & "/" & sHi
        ElseIf sLo <> "" And sHi = "" Then
          .lblSpec_MER(iBand).Caption = ">= " & sLo
        ElseIf sLo = "" And sHi <> "" Then
          .lblSpec_MER(iBand).Caption = "<= " & sHi
        Else
          'No Spec ????
          .lblSpec_MER(iBand).Caption = "No spec?"
        End If
      Else
        .lblBER_MHz(iBand).Visible = False
        .lblSpec_BER(iBand).Visible = False
        .lblSpec_MER(iBand).Visible = False
        '***
        .txtMeas_BER(iBand).Visible = False
        .txtMeas_MER(iBand).Visible = False
        '***
        .ProgressBar1(iBand).Visible = False
      End If
    Next iBand
  End With
  
  
  '---------- Test Model Combo
  
  'Connection string
  sConn = "Provider=Microsoft.Jet.OLEDB.4.0"
  sConn = sConn & ";Data Source=" & gPATH.sConfig & "\Spec_BER.mdb"
  sConn = sConn & ";Persist Security Info=False"
  
  'SQL string
  sSQL = ""
  'sSQL = sSQL & "SELECT * FROM [Spec] "
  sSQL = sSQL & "SELECT * FROM [" & gSPEC.sWL_Table(giIndex) & "] "
  sSQL = sSQL & "ORDER BY [Test_Model] ASC "
  'Open Record
  rs.Open sSQL, sConn, adOpenKeyset, adLockReadOnly

  If rs.EOF = True Then
    MsgBox "Wrong Model??? --- No tests!!!"
    End
  End If
  
  iMax = rs.RecordCount
  
  '***** Test Type
  Me.cboTestModel.Clear
  For iCnt = 1 To iMax
    
    Me.cboTestModel.AddItem rs("Test_Model")
  
    If iCnt < iMax Then
      rs.MoveNext
    End If
  
  Next iCnt
  
  Me.cboTestModel.ListIndex = 0
  
  
End Sub



Private Sub cboTestModel_Click()

  If Me.cboTestModel.Text = "" Then
    Exit Sub
  End If
  
  Call mSQL.LoadWLPwr

End Sub


Private Sub cmdClear_Click()
  
  Dim iBand As Integer
  
  For iBand = 1 To gconMAX_BAND
    Me.txtMeas_BER(iBand).Text = ""
    Me.txtMeas_BER(iBand).BackColor = Pass_Color
    '***
    Me.txtMeas_MER(iBand).Text = ""
    Me.txtMeas_MER(iBand).BackColor = Pass_Color
    '***
    Me.ProgressBar1(iBand).value = 0
  Next iBand
  
End Sub

Private Sub cmdRecall_Click()
  
  Me.txtSN.Text = Trim(Me.txtSN.Text)
  If Me.txtSN = "" Then
    MsgBox "Please enter S/N before recalling test data"
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  Me.txtStatus.Text = "--- Recalling Data --- Please wait"
  Call mSQL.RecallData
  
  Screen.MousePointer = vbNormal
  
  Me.txtStatus.Text = "--- Done ---"
  
End Sub


Private Sub cmdRUN_Click()
  
  Dim iCnt As Integer
  
  Dim dMax As Double
  
  Dim sBuf As String
  
  Dim dBER As Double
  Dim dMER As Double
  Dim dOffset As Double
  
  Dim sMeasBER As String
  Dim sMeasMER As String
  
  Dim sResult As String
  
  Dim sMeas As String
  
  
  '***** S/N???
  Me.txtSN.Text = Trim(Me.txtSN.Text)
  If Me.txtSN.Text = "" Then
    MsgBox "Please enter S/N before proceeding"
    Me.txtSN.SetFocus
    Exit Sub
  End If
  
  '***** Operator???
  Me.txtOperator.Text = Trim(Me.txtOperator.Text)
  If Me.txtOperator.Text = "" Then
    MsgBox "Please enter Operator Name before proceeding"
    Me.txtOperator.SetFocus
    Exit Sub
  End If
  
  '***** WL Meas???
  If Trim(Me.txtWL.Text) = "" Then
    MsgBox "Please enter WL in nm before proceeding"
    Me.txtWL.SetFocus
    Exit Sub
  Else
    Me.txtWL.Text = Trim(Me.txtWL.Text)
    If Me.lblSpec_WL.Caption <> "" Then
      '***** Check limits
      sBuf = Me.lblSpec_WL.Caption
      sMeas = Me.txtWL.Text
      Call mGeneral.CheckLimits(sBuf, sMeas, sResult)
      If Right(sResult, 1) = "*" Then
        MsgBox "WL entry does not meet spec --- Try again"
        Me.txtWL.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  '***** Pwr dBm Meas???
  If Trim(Me.txtPwrDBM.Text) = "" Then
    MsgBox "Please enter Opt Pwr in dBm before proceeding"
    Me.txtPwrDBM.SetFocus
    Exit Sub
  Else
    Me.txtPwrDBM.Text = Trim(Me.txtPwrDBM.Text)
    If Me.lblSpec_Pwr.Caption <> "" Then
      '***** Check limits
      sBuf = Me.lblSpec_Pwr.Caption
      sMeas = Me.txtPwrDBM.Text
      Call mGeneral.CheckLimits(sBuf, sMeas, sResult)
      If Right(sResult, 1) = "*" Then
        MsgBox "Pwr dBm entry does not meet spec --- Try again"
        Me.txtPwrDBM.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  
  
  gbPass = True
  
  gbStop = False
  
  
  
  '***** Init Progressive bars
  For iCnt = 1 To gconMAX_BAND
    Me.ProgressBar1(iCnt).value = 0
    '****
    Me.txtMeas_BER(iCnt).Text = ""
    Me.txtMeas_BER(iCnt).BackColor = Pass_Color
    '****
    Me.txtMeas_MER(iCnt).Text = ""
    Me.txtMeas_MER(iCnt).BackColor = Pass_Color
  Next iCnt
  
  '------------ Dis/Enable buttons
  Me.cmdClear.Enabled = False
  Me.cmdRecall.Enabled = False
  Me.cmdRUN.Enabled = False
  Me.cmdStop.Enabled = True
  
  
  Me.fraInfo.Enabled = False
  
  '------------ Set Display to History BER
  Call gEQ.clsBER.SetQAMHistory_Display
  
  
  For iCnt = 1 To gconMAX_BAND
    
    If Me.lblBER_MHz(iCnt).Visible = True Then
      
      '****** Offset
      If gSPEC.sOffset(iCnt, giIndex) = "" Then
        dOffset = 0
      Else
        dOffset = CDbl(gSPEC.sOffset(iCnt, giIndex))
      End If
      
      dMax = gSPEC.dGate_Sec(giIndex)
      
      sBuf = "Testing BER/MER --- " & Me.lblBER_MHz(iCnt) & " (Needs 13 Sec setup)"
      Me.txtStatus.Text = sBuf
      
      giTimerCnt = 0
      
      giCurrBand = 1
      
      Me.ProgressBar1(iCnt).Max = dMax
      
      Screen.MousePointer = vbHourglass
      
      '***** BER Set up
      Call gEQ.clsBER.SetRFFrequencyMHz(Val(Me.lblBER_MHz(iCnt).Caption))
      'delay (1)
      Call gEQ.clsBER.SetQAMHistory_Restart
      delay (13)
      
      Screen.MousePointer = vbNormal
      
      Do
                
                '========= Allow interruption
                If gbStop = True Then
                  '---- Dis/Enable buttons
                  Me.fraInfo.Enabled = True
                  '***
                  Me.cmdClear.Enabled = True
                  Me.cmdRecall.Enabled = True
                  Me.cmdRUN.Enabled = True
                  Me.cmdStop.Enabled = False
                  '***
                  Me.txtStatus.Text = "--- Stop Tests ---"
                  '*** Reset flag
                  gbStop = False
                  '***
                  Exit Sub
                End If
        
        
        '*** Looping
        giTimerCnt = giTimerCnt + 1
        Me.ProgressBar1(iCnt).value = giTimerCnt
        delay (1)
        '**** Current BER/MER Reads
        Call gEQ.clsBER.ReadQAMHistBER_Before_Curr(dBER)
        Call gEQ.clsBER.ReadQAMHistMER_DB_Curr(dMER)
        'Call gEQ.clsBER.ReadQAMHistBER_Before_Ave(dBER)
        'Call gEQ.clsBER.ReadQAMHistMER_DB_Ave(dMER)
        '*** Display
        Me.txtStatus.Text = sBuf & vbCrLf
        Me.txtStatus.Text = Me.txtStatus.Text & Format(dBER, "0.00e+000") & Space(5) & dMER
      Loop Until giTimerCnt >= dMax
  
      '***** Test results (history BER/MER Avg
      Call gEQ.clsBER.ReadQAMHistBER_Before_Ave(dBER)
      sMeasBER = Format(dBER, "0.00e+000")
      Call CheckLimits(Me.lblSpec_BER(iCnt), sMeasBER, sResult)
      'Me.txtMeas_BER(iCnt).Text = sResult
      Me.txtMeas_BER(iCnt).Text = sMeasBER
      If Right(sResult, 1) <> "*" Then
        Me.txtMeas_BER(iCnt).BackColor = Pass_Color
      Else
        Me.txtMeas_BER(iCnt).BackColor = Fail_Color
        gbPass = False
      End If
      '***
      Call gEQ.clsBER.ReadQAMHistMER_DB_Ave(dMER)
      
      sMeasMER = dMER + dOffset
 
      
      Call CheckLimits(Me.lblSpec_MER(iCnt), sMeasMER, sResult)
      'Me.txtMeas_MER(iCnt).Text = sResult
      Me.txtMeas_MER(iCnt).Text = sMeasMER
      If Right(sResult, 1) <> "*" Then
        Me.txtMeas_MER(iCnt).BackColor = Pass_Color
      Else
        Me.txtMeas_MER(iCnt).BackColor = Fail_Color
        gbPass = False
      End If
      
    End If
    
  Next iCnt
  
  
  '--------- Save Data
  Screen.MousePointer = vbHourglass
  Me.txtStatus.Text = "--- Saving Data ---"
  Call mSQL.SaveData
  Screen.MousePointer = vbNormal
  
  
  Me.fraInfo.Enabled = True
  
  Me.txtStatus.Text = "--- Done ---"
  
  
  '------------ Enable buttons
  Me.cmdClear.Enabled = True
  Me.cmdRecall.Enabled = True
  Me.cmdRUN.Enabled = True
  Me.cmdStop.Enabled = False
  
  
End Sub

Private Sub cmdStop_Click()
  
  gbStop = True
  
End Sub


Private Sub Form_Load()
  
  Dim iCnt As Integer
  
  Dim sBuf As String
  Dim lSize As Long
  
  Dim sArray() As String
  
  
  '===== Define Path Names here
  ChDir (App.Path)    'Set directory to Application Path
  gPATH.sApp = App.Path
  gPATH.sCurDir = App.Path
  If UCase(gPATH.sCurDir) <> "C:\" Then
    ChDir ("..")      'Up 1 level
    gPATH.sCurDir = CurDir
  Else
    MsgBox "Fatal Error!!! ---- Wrong Path --- Call Test Engineer"
    End
  End If
  gPATH.sConfig = gPATH.sCurDir & "\Config_BER"
  gPATH.sTestData = gPATH.sCurDir & "\TestData"
  '===== Must use C: root to store test data
  '=====  for Birth Certificate creation with DOS Batch
  '=====  file conv.bat pointing to C: root folder
  'gPATH.sTestData = "C:"    'Use C: root to store test data folders
  ChDir (gPATH.sApp)      'Set Curr Dir to Application Dir
  '***** Computer name
  sBuf = Space(100)
  lSize = Len(sBuf)
  Call GetComputerName(sBuf, lSize)
  gPATH.sComputerName = Left(sBuf, lSize)
  '==================================================
  
  
  '========== GPIB Addr????
  Call mGeneral.FileToKeyArray(gPATH.sConfig & "\GPIBAddr.txt", sArray())
  '***** RF ATT --- LW8200
  gGPIB.iBER = 6
  sBuf = mGeneral.ReadKeyArray(sArray(), "BER_Rx", "GPIB_Addr")
  If Val(sBuf) > 0 Then
    gGPIB.iBER = Val(sBuf)
  End If
  
  '------------------- GPIB Init here
  gblnEQInit = False
  Call mEquip.GPIBInit
  '***** Init Successful?
  If gblnEQInit = False Then
    MsgBox "Failed GPIB init!!!"
  End If
  
  
  '------------------- SQL Table
  If Dir(gPATH.sConfig & "\SQLSetup.txt") = "" Then
    Call mSQL.InitSQLDBase
  End If
  Call FileToKeyArray(gPATH.sConfig & "\SQLSetup.txt", gsKeyDBase)
  gDB.SQLServer = ReadKeyArray(gsKeyDBase, "SQL", "Server")
  gDB.SQLDatabase = ReadKeyArray(gsKeyDBase, "SQL", "Database")
  gDB.SQLPassword = ReadKeyArray(gsKeyDBase, "SQL", "Password")
  gDB.SQLUser = ReadKeyArray(gsKeyDBase, "SQL", "User")
  
  
  '------------------- Title
  Call mGeneral.RevisionHistory
  Me.Caption = "CATV BER Test" & Space(5) & gSID & _
                Space(5) & "(" & gDB.SQLServer & ")"
  
  
  '------------------- Load Test Location etc...
  mGeneral.LoadTestLoc
  '***
  Me.txtStationID = gSTA.sStationID
  Me.txtTestLoc = gSTA.sTestLoc
  Me.txtDBaseType = gSTA.sDBaseType
  
  
  '========== Load Specs
  Call mSQL.LoadSpec
  
  '***** Main Model
  Me.cboModel.Clear
  Me.cboModel.AddItem "Select"
  For iCnt = 1 To UBound(gSPEC.sModel)
    Me.cboModel.AddItem gSPEC.sModel(iCnt)
  Next iCnt
  Me.cboModel.ListIndex = 0
  giIndex = 0
  
  '***** Test Type
  Me.cboTestType.Clear
  For iCnt = 1 To UBound(gTEST_TYPE.sTest_Type)
    Me.cboTestType.AddItem gTEST_TYPE.sTest_Type(iCnt)
  Next iCnt
  '***** Show PROD?
  Me.cboTestType.ListIndex = 0
  For iCnt = 0 To (Me.cboTestType.ListCount - 1)
    If UCase(Me.cboTestType.List(iCnt)) = "PROD" Then
      Me.cboTestType.ListIndex = iCnt
      Exit For
    End If
  Next iCnt
  
  Me.cmdStop.Enabled = False
  
  'Call mGeneral.FileToKeyArray(gPATH.sConfig & "\Offset.txt", sArray())
  
  'gOffset = mGeneral.ReadKeyArray(sArray(), "Offset", "Offset_991")
    
End Sub


Private Sub mnuGPIBErrorDisp_Click()

  If mnuGPIBErrorDisp.Checked Then
    Call gEQ.clsBER.DisplayMsgOFF
    mnuGPIBErrorDisp.Checked = False
  Else
    Call gEQ.clsBER.DisplayMsgON
    mnuGPIBErrorDisp.Checked = True
  End If

End Sub


Private Sub Timer1_Timer()
  
  '***** Display Date/Time
  Me.txtDate.Text = Date
  Me.txtTime.Text = Time
  
End Sub


Private Sub txtLotN_KeyPress(KeyAscii As Integer)

  'If KeyAscii = 13 Then   'if hit enter
  If KeyAscii = vbKeyReturn Then   'if hit enter
    Me.txtLotN.Text = Trim(Me.txtLotN.Text)
    Me.txtWO.SetFocus
  End If

End Sub


Private Sub txtOperator_KeyPress(KeyAscii As Integer)

  'If KeyAscii = 13 Then   'if hit enter
  If KeyAscii = vbKeyReturn Then   'if hit enter
    Me.txtOperator.Text = Trim(Me.txtOperator.Text)
    Me.txtLotN.SetFocus
  End If

End Sub


Private Sub txtPwrDBM_KeyPress(KeyAscii As Integer)

  Dim sBuf As String
  Dim sMeas As String
  Dim sResult As String
  

  'If KeyAscii = 13 Then   'if hit enter
  If KeyAscii = vbKeyReturn Then   'if hit enter
    Me.txtPwrDBM.Text = Trim(Me.txtPwrDBM.Text)
    If Me.lblSpec_Pwr.Caption <> "" Then
      '***** Check limits
      sBuf = Me.lblSpec_Pwr.Caption
      sMeas = Me.txtPwrDBM.Text
      Call mGeneral.CheckLimits(sBuf, sMeas, sResult)
      If Right(sResult, 1) = "*" Then
        MsgBox "Opt Pwr entry does not meet spec --- Try again"
        Me.txtPwrDBM.SetFocus
        Exit Sub
      End If
    End If
    Me.cmdRUN.SetFocus
  End If

End Sub


Private Sub txtSN_KeyPress(KeyAscii As Integer)
  
  'If KeyAscii = 13 Then   'if hit enter
  If KeyAscii = vbKeyReturn Then   'if hit enter
    Me.txtSN.Text = Trim(Me.txtSN.Text)
    Me.txtOperator.SetFocus
  End If
  
End Sub


Private Sub txtWL_KeyPress(KeyAscii As Integer)

  Dim sBuf As String
  Dim sMeas As String
  Dim sResult As String
  
  
  'If KeyAscii = 13 Then   'if hit enter
  If KeyAscii = vbKeyReturn Then   'if hit enter
    Me.txtWL.Text = Trim(Me.txtWL.Text)
    If Me.lblSpec_WL.Caption <> "" Then
      '***** Check limits
      sBuf = Me.lblSpec_WL.Caption
      sMeas = Me.txtWL.Text
      Call mGeneral.CheckLimits(sBuf, sMeas, sResult)
      If Right(sResult, 1) = "*" Then
        MsgBox "WL entry does not meet spec --- Try again"
        Me.txtWL.SetFocus
        Exit Sub
      End If
    End If
    Me.txtPwrDBM.SetFocus
  End If

End Sub


Private Sub txtWO_KeyPress(KeyAscii As Integer)

  'If KeyAscii = 13 Then   'if hit enter
  If KeyAscii = vbKeyReturn Then   'if hit enter
    Me.txtWO.Text = Trim(Me.txtWO.Text)
    Me.txtWL.SetFocus
  End If

End Sub


