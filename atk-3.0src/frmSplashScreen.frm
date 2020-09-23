VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplashScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Attack Tool Kit"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmSplashScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Arrow and Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timTimer 
      Interval        =   55
      Left            =   4080
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar pbrStatus 
      Height          =   135
      Left            =   1800
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStatusInformation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "loading the software into the memory"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attack Tool Kit is starting ... Please wait!"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   480
      Width           =   4695
   End
   Begin VB.Shape shpRedLine 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Caption = SoftwareName & " starting ..."
    lblStatus.Caption = SoftwareName & " is starting ... Please wait!"
End Sub

Private Sub timTimer_Timer()
    timTimer.Enabled = False
    Load frmMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSplashScreen = Nothing
End Sub
