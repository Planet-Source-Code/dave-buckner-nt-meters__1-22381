VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000001&
   Caption         =   "ARRiVE's Meter Tester"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6285
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VertMenuTester.VMeterWideLG VMeterWideLG1 
      Height          =   3360
      Left            =   3480
      TabIndex        =   13
      Top             =   600
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   5927
      Value           =   0
   End
   Begin VertMenuTester.VMeterWideSM VMeterWideSM1 
      Height          =   1935
      Left            =   4080
      TabIndex        =   12
      Top             =   2040
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   3413
      Value           =   0
   End
   Begin VertMenuTester.VMeterNarrowLG VMeterNarrowLG1 
      Height          =   3360
      Left            =   4680
      TabIndex        =   11
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5927
   End
   Begin VertMenuTester.VMeterNarrowSM VMeterNarrowSM1 
      Height          =   1935
      Left            =   5160
      TabIndex        =   10
      Top             =   2040
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   3413
      Value           =   0
      BackColor       =   8388736
      ForeColor       =   16711935
   End
   Begin VertMenuTester.VDualMeterSM VDualMeterSM2 
      Height          =   1935
      Left            =   2640
      TabIndex        =   9
      Top             =   2040
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   5927
      Value           =   0
      BackColor       =   16744703
      ForeColor       =   8388736
      BorderColor     =   -2147483647
   End
   Begin VertMenuTester.VDualMeterLG VDualMeterLG2 
      Height          =   3360
      Left            =   1800
      TabIndex        =   8
      Top             =   600
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   5927
      BackColor       =   4194368
      ForeColor       =   16711935
      BorderColor     =   -2147483647
   End
   Begin VertMenuTester.VDualMeterSM VDualMeterSM1 
      Height          =   1935
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   5927
      Value           =   0
      BackColor       =   65280
      ForeColor       =   16384
   End
   Begin VertMenuTester.VDualMeterLG VDualMeterLG1 
      Height          =   3360
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   5927
   End
   Begin VertMenuTester.HMeterTallLG HMeterTallLG1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   714
      Value           =   0
      BackColor       =   16776960
      ForeColor       =   8421376
   End
   Begin VertMenuTester.HMeterTallSM HMeterTallSM1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   714
      Value           =   0
      BackColor       =   65280
      ForeColor       =   16384
      BorderColor     =   -2147483647
   End
   Begin VertMenuTester.HMeterSlimLG HMeterSlimLG1 
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   714
      Value           =   0
      BackColor       =   8421376
      ForeColor       =   16776960
   End
   Begin VertMenuTester.HMeterSlimSM HMeterSlimSM1 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   714
      Value           =   0
      BorderColor     =   -2147483647
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   960
      Max             =   1000
      Min             =   10
      TabIndex        =   0
      Top             =   240
      Value           =   10
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin VB.Label lblInterval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Speed = 10"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.Menu Menu 
      Caption         =   "Start"
      Index           =   0
   End
   Begin VB.Menu Menu 
      Caption         =   "Stop"
      Index           =   1
   End
   Begin VB.Menu Menu 
      Caption         =   "Multiplier"
      Index           =   2
      Begin VB.Menu mnuMultiplier 
         Caption         =   "x1"
         Index           =   0
      End
      Begin VB.Menu mnuMultiplier 
         Caption         =   "x2"
         Index           =   1
      End
      Begin VB.Menu mnuMultiplier 
         Caption         =   "x10"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngCount As Long
Dim lngPlus As Long
Private Sub Form_Load()
    VDualMeterLG1.ReDraw
    VDualMeterLG2.ReDraw

    VDualMeterSM1.ReDraw
    VDualMeterSM2.ReDraw
    
    VMeterWideLG1.ReDraw
    VMeterWideSM1.ReDraw
    
    VMeterNarrowSM1.ReDraw
    VMeterNarrowLG1.ReDraw
    
    HMeterSlimSM1.ReDraw
    HMeterSlimLG1.ReDraw
    
    HMeterTallLG1.ReDraw
    HMeterTallSM1.ReDraw
    
    Timer1.Enabled = False
    lngPlus = 1
End Sub
Private Sub HScroll1_Change()
    Timer1.Interval = HScroll1.Value
    lblInterval.Caption = "Current Speed = " & HScroll1.Value
End Sub
Private Sub Menu_Click(Index As Integer)
    Select Case Index
        Case 0  'Start
            Timer1.Enabled = True
            lngCount = 0
        Case 1  'Stop
            Timer1.Enabled = False
    End Select
End Sub
Private Sub mnuMultiplier_Click(Index As Integer)
    Select Case Index
        Case 0  'x1
            lngPlus = 1
            lngCount = 0
        Case 1  'x2
            lngPlus = 2
            lngCount = 0
        Case 2  'x10
            lngPlus = 10
            lngCount = 0
    End Select
End Sub
Private Sub Timer1_Timer()
    lngCount = lngCount + lngPlus
    VDualMeterLG1.Value = lngCount
    VDualMeterLG2.Value = lngCount
    
    VDualMeterSM1.Value = lngCount
    VDualMeterSM2.Value = lngCount
    
    VMeterWideLG1.Value = lngCount
    VMeterWideSM1.Value = lngCount
    
    VMeterNarrowSM1.Value = lngCount
    VMeterNarrowLG1.Value = lngCount
    
    HMeterSlimLG1.Value = lngCount
    HMeterSlimSM1.Value = lngCount
    
    HMeterTallLG1.Value = lngCount
    HMeterTallSM1.Value = lngCount
    If lngCount = 100 Then
        lngCount = 0
        'Timer1.Enabled = False
    End If
End Sub
