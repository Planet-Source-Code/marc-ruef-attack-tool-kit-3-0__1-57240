VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   4230
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7095
   Icon            =   "frmConfiguration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSpeech 
      Caption         =   "Speech"
      Height          =   3375
      Left            =   240
      TabIndex        =   40
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkActivateSpeech 
         Caption         =   "Activate Speech"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label lblTestSpeechFeature 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Test the speech feature."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2400
         MouseIcon       =   "frmConfiguration.frx":0CCA
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Top             =   1680
         Width           =   1755
      End
      Begin VB.Label lblSpeechDescription 
         Caption         =   $"frmConfiguration.frx":0FD4
         Height          =   615
         Left            =   480
         TabIndex        =   41
         Top             =   720
         Width           =   5895
      End
   End
   Begin VB.Frame fraSearchengine 
      Caption         =   "Searchengine"
      Height          =   3375
      Left            =   240
      TabIndex        =   56
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cmbSearchEngineURL 
         Height          =   315
         Left            =   240
         TabIndex        =   57
         Text            =   "http://www.google.com/search?q="
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label lblSearchEngineTest 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Test the selected search engine query url by searching for the ATK project."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   645
         MouseIcon       =   "frmConfiguration.frx":10C8
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   1440
         Width           =   5325
      End
      Begin VB.Label lblSearchEngineURLName 
         Caption         =   "Search engine default query string for online searches"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame fraTarget 
      Caption         =   "Target"
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   6615
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   120
         MaxLength       =   200
         TabIndex        =   1
         Text            =   "localhost"
         ToolTipText     =   "Host name or IP address of the target"
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label lblDisclaimer 
         Alignment       =   2  'Center
         Caption         =   "Warning: You should never scan a network ressource without permission."
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   1560
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   7200
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame fraHelp 
      Caption         =   "Help"
      Height          =   3375
      Left            =   240
      TabIndex        =   52
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cmbHelpURL 
         Height          =   315
         Left            =   1080
         TabIndex        =   53
         Text            =   "http://www.computec.ch/projekte/atk/documentation/help/"
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label lblOnlineHelpURLTest 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Test the selected online help url."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2145
         MouseIcon       =   "frmConfiguration.frx":13D2
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   1920
         Width           =   2325
      End
      Begin VB.Label lblHelpDescription 
         Caption         =   $"frmConfiguration.frx":16DC
         Height          =   615
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label lblHelpURLName 
         Caption         =   "Help URL"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame fraAlerting 
      Caption         =   "Alerting"
      Height          =   3375
      Left            =   240
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkAlertingVulnerabilityNotFound 
         Caption         =   "Produce alert when vulnerbility is not found."
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   720
         Width           =   3495
      End
      Begin VB.CheckBox chkAlertingVulnerabilityFound 
         Caption         =   "Produce alert when vulnerability is found."
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame fraLogs 
      Caption         =   "Logs"
      Height          =   3375
      Left            =   240
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cmbLogsSecurityLevel 
         Height          =   315
         ItemData        =   "frmConfiguration.frx":1797
         Left            =   1320
         List            =   "frmConfiguration.frx":1799
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1560
         Width           =   5175
      End
      Begin VB.DirListBox dirLogs 
         Height          =   1215
         Left            =   1320
         TabIndex        =   43
         Top             =   2040
         Width           =   5175
      End
      Begin VB.CheckBox chkActivateLogs 
         Caption         =   "Activate lo&gs"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Value           =   1  'Checked
         Width           =   6255
      End
      Begin VB.Label lblLogSecurityLevel 
         Caption         =   "Security Level"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblLogsDirectory 
         Caption         =   "Logs directory"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   $"frmConfiguration.frx":179B
         Height          =   615
         Left            =   360
         TabIndex        =   34
         Top             =   720
         Width           =   6015
      End
   End
   Begin VB.Frame fraMapping 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   240
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame fraICMPMapping 
         Caption         =   "ICMP Mapping"
         Height          =   3375
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   6615
         Begin VB.CheckBox chkDoICMPMapping 
            Caption         =   "Do &ICMP mapping (ICMP echo request)"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Value           =   1  'Checked
            Width           =   6375
         End
         Begin VB.CheckBox chkScanifICMPfails 
            Caption         =   "Scan if ICMP mapping fails"
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   720
            Width           =   2295
         End
      End
   End
   Begin MSComDlg.CommonDialog cdgSaveAs 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save file"
      FileName        =   "newdefault.config"
      Filter          =   "ATK Configuration (*.config)|*.config|INI Files (*.ini)|*.ini|All Files (*.*)|*.*"
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reporting"
      Height          =   3375
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.DirListBox dirReportsDirectory 
         Height          =   2565
         Left            =   1560
         TabIndex        =   36
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label lblReportTemplateNote 
         Caption         =   "Editing of the report templates can be done in the report configuration."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   3000
         Width           =   4935
      End
      Begin VB.Label lblReportsDirectory 
         Caption         =   "Reports Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open file"
      FileName        =   "default.config"
      Filter          =   "ATK Configuration (*.config)|*.config|INI Files (*.ini)|*.ini|All Files (*.*)|*.*"
   End
   Begin VB.Frame fraPlugins 
      Caption         =   "Plugins"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cmbPluginsDownloadURL 
         Height          =   315
         Left            =   1560
         TabIndex        =   51
         Text            =   "http://www.computec.ch/projekte/atk/plugins/pluginslist/"
         Top             =   1680
         Width           =   4935
      End
      Begin VB.DirListBox dirPlugins 
         Height          =   1215
         Left            =   1560
         TabIndex        =   20
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtDefaultSleep 
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   23
         Text            =   "3000"
         ToolTipText     =   "Default wait time for sleep command"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtTimeout 
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   22
         Text            =   "30000"
         ToolTipText     =   "Timeout for the plugins"
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblPluginsDownloadURL 
         Caption         =   "Plugins Download"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblSleepValueDefault 
         Caption         =   "(Default: 3000 = 3 seconds)"
         Height          =   255
         Left            =   2400
         TabIndex        =   38
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label lblSleepValueName 
         Caption         =   "Default wait value (ms) for sleep command"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblTimeoutDefault 
         Caption         =   "(Default: 30000 = 30 seconds)"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblPluginsDirectoryName 
         Caption         =   "Plugins Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblTimeoutName 
         Caption         =   "Timeout (ms) for stucked plugins"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.Frame fraPreferences 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame fraSafety 
         Caption         =   "Safety"
         Height          =   2175
         Left            =   0
         TabIndex        =   25
         Top             =   1200
         Width           =   6615
         Begin VB.CheckBox chkDoSilentChecks 
            Caption         =   "&Do silent checks"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkDoNoDoSChecks 
            Caption         =   "Do no Denial of Service checks"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label lblDoSilentChecksDescription 
            Caption         =   $"frmConfiguration.frx":1876
            Height          =   495
            Left            =   480
            TabIndex        =   32
            Top             =   720
            Width           =   6015
         End
         Begin VB.Label lblDonoDoSDescription 
            Caption         =   $"frmConfiguration.frx":191C
            Height          =   495
            Left            =   480
            TabIndex        =   31
            Top             =   1560
            Width           =   6015
         End
      End
      Begin VB.Frame fraMode 
         Caption         =   "Mode"
         Height          =   1095
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   6615
         Begin VB.OptionButton optSingleCheck 
            Caption         =   "&Single Check"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Only check specific potential flaws on demand."
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optFullAudit 
            Caption         =   "&Full Audit"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Check the target for all possible potential flaws."
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblSingleCheckDescription 
            Caption         =   "Only check specific potential flaws on demand."
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label lblFullAuditDescription 
            Caption         =   "Check the target for all possible potential flaws."
            Height          =   255
            Left            =   2040
            TabIndex        =   21
            Top             =   720
            Width           =   4455
         End
      End
   End
   Begin VB.Frame fraSuggestions 
      Caption         =   "Suggestions"
      Height          =   3375
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkSuggestions 
         Caption         =   "&Activate suggestions"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.DirListBox dirSuggestions 
         Height          =   2565
         Left            =   1800
         TabIndex        =   28
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblSuggestionsDirectory 
         Caption         =   "Suggestions Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame fraInteractivity 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   240
      TabIndex        =   39
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
   End
   Begin MSComctlLib.TabStrip tspConfiguration 
      Height          =   3975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7011
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   11
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Target"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "P&references"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mapping"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Plugins"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Alerting"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Suggestions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search&engine"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reporting"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Logs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Speech"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenItem 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSpererator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveItem 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAsItem 
         Caption         =   "Save As ..."
      End
      Begin VB.Menu mnuFileSpererator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseItem 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpConfigurationHelpItem 
         Caption         =   "&Configuration Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 3.0 2004-11-05                                                           *
' * - Added an error routine withing the save as function if the cancel button is    *
' *   pressed.                                                                       *
' * Version 3.0 2004-11-03                                                           *
' * - Fixed a bug if opening the report configuration.                               *
' * Version 3.0 2004-11-01                                                           *
' * - Fixed the File/New function. It should work now.                               *
' * - Deleted all not needed nor supported elements.                                 *
' * Version 3.0 2004-10-30                                                           *
' * - Fully enhanced and re-sorted the configuration file output. We are now using a *
' *   Unix/Linux conf file format that allows commenting out lines by using the #    *
' * Version 3.0 2004-10-28                                                           *
' * - Added the menu save as function to save specific/other configuration files.    *
' * Version 3.0 2004-10-23                                                           *
' * - Added the menu file open function to open specific/other configuration files.  *
' * Version 3.0 2004-10-22                                                           *
' * - Added the error message behavior if the target specifying is wrong.            *
' * Version 3.0 2004-10-20                                                           *
' * - Added the tab and routines for the online help configuration.                  *
' * Version 3.0 2004-10-08                                                           *
' * - Enhanced and bugfixed the whole logging.                                       *
' * - Added the update features for the AutoUpdate.                                  *
' * Version 2.1 2004-09-08                                                           *
' * - Corrected and enhanced the full audit mode warning.                            *
' * Version 2.0 2004-04-08                                                           *
' * - Added the actualizing of the target data in frmAttackVisualizing after         *
' *   clicking accept.                                                               *
' * Version 1.1 2004-03-20                                                           *
' * - Added the configuration file name in the frame caption for more verbosity.     *
' * - Added a warning message if the full audit mode is selected.                    *
' ************************************************************************************

Private Sub chkActivateLogs_Click()
    If chkActivateLogs.Value = 1 Then
        cmbLogsSecurityLevel.Enabled = True
        dirLogs.Enabled = True
    Else
        cmbLogsSecurityLevel.Enabled = False
        dirLogs.Enabled = False
    End If
End Sub

Private Sub chkDoICMPMapping_Click()
    If chkDoICMPMapping.Value = 0 Then
        chkScanifICMPfails.Enabled = False
    Else
        chkScanifICMPfails.Enabled = True
    End If
End Sub

Private Sub chkSuggestions_Click()
    If chkSuggestions.Value <> 1 Then
        dirSuggestions.Enabled = False
    Else
        dirSuggestions.Enabled = True
    End If
End Sub

Private Sub LoadActualConfigurationValues()
    'Show the frame title with configuration file name
    frmConfiguration.Caption = "Configuration - " & application_configuration_filename
    
    'Display and activate the loaded config data
    txtTarget.Text = Target
    txtTimeout.Text = AttackTimeout
    txtDefaultSleep.Text = DefaultSleepValue
    cmbPluginsDownloadURL.Text = PluginsDownloadURL
    cmbSearchEngineURL.Text = SearchEngineURL
    cmbHelpURL.Text = HelpURL
    
    If (Dir$(PluginDirectory, 16) <> "") Then
        dirPlugins.Path = PluginDirectory
    End If
    
    If (Dir$(SuggestionsDirectory, 16) <> "") Then
        dirSuggestions.Path = SuggestionsDirectory
    End If

    'dirReportsDirectory.Path = ReportsDirectory
    
    If AttackMode = "SingleCheck" Then
        optSingleCheck.Value = True
        optFullAudit.Value = False
    ElseIf AttackMode = "FullAudit" Then
        optSingleCheck.Value = False
        optFullAudit.Value = True
    End If

    If DoSilentChecks = True Then
        chkDoSilentChecks.Value = 1
    Else
        chkDoSilentChecks.Value = 0
    End If
    
    If DoNoDoSChecks = True Then
        chkDoNoDoSChecks.Value = 1
    Else
        chkDoNoDoSChecks.Value = 0
    End If

    If DoICMPMapping = True Then
        chkDoICMPMapping.Value = 1
    Else
        chkDoICMPMapping.Value = 0
    End If

    If ScanIfICMPFails = True Then
        chkScanifICMPfails.Value = 1
    Else
        chkScanifICMPfails.Value = 0
    End If
    
    If AlertingVulnFound = True Then
        chkAlertingVulnerabilityFound.Value = 1
    Else
        chkAlertingVulnerabilityFound.Value = 0
    End If
    
    If AlertingVulnNotFound = True Then
        chkAlertingVulnerabilityNotFound.Value = 1
    Else
        chkAlertingVulnerabilityNotFound.Value = 0
    End If
    
    If ActivateSuggestions = True Then
        chkSuggestions.Value = 1
    Else
        chkSuggestions.Value = 0
    End If
    
    If ActivateLogs = True Then
        chkActivateLogs.Value = 1
    Else
        chkActivateLogs.Value = 0
    End If
    
    On Error Resume Next 'Workaround!
    dirLogs.Path = LogsDirectory
    
    If ActivateSpeech = True Then
        chkActivateSpeech.Value = 1
    Else
        chkActivateSpeech.Value = 0
    End If

    If LogsSecurityLevel = 0 Then
        cmbLogsSecurityLevel.ListIndex = 0
    ElseIf LogsSecurityLevel = 1 Then
        cmbLogsSecurityLevel.ListIndex = 1
    ElseIf LogsSecurityLevel = 2 Then
        cmbLogsSecurityLevel.ListIndex = 2
    ElseIf LogsSecurityLevel = 3 Then
        cmbLogsSecurityLevel.ListIndex = 3
    ElseIf LogsSecurityLevel = 4 Then
        cmbLogsSecurityLevel.ListIndex = 4
    ElseIf LogsSecurityLevel = 5 Then
        cmbLogsSecurityLevel.ListIndex = 5
    ElseIf LogsSecurityLevel = 6 Then
        cmbLogsSecurityLevel.ListIndex = 6
    Else
        cmbLogsSecurityLevel.ListIndex = 7
    End If
End Sub

Private Sub cmbSearchEngineURL_KeyPress(KeyAscii As Integer)
    'Complete a combobox writing
    Static iLeftOff As Long
    ComboAutoComplete cmbSearchEngineURL, KeyAscii, iLeftOff
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Show the frame title with configuration file name
    frmConfiguration.Caption = "Configuration - " & application_configuration_filename
      
    'Add some default values
    cmbLogsSecurityLevel.AddItem "0 emergencies (A panic condition if the system is unusable.)", 0
    cmbLogsSecurityLevel.AddItem "1 alerts (A condition that should be corrected immediately.)", 1
    cmbLogsSecurityLevel.AddItem "2 critical (Critical conditions, e.g. hard device errors.)", 2
    cmbLogsSecurityLevel.AddItem "3 error (Errors)", 3
    cmbLogsSecurityLevel.AddItem "4 warnings (Warning messages)", 4
    cmbLogsSecurityLevel.AddItem "5 notifications (Conditions that are not error conditions, but should possibly be handled specially.)", 5
    cmbLogsSecurityLevel.AddItem "6 informational (Informational messages)", 6
    cmbLogsSecurityLevel.AddItem "7 debugging (Messages that contain information normally of use only when debugging a program.)", 7
    
    cmbPluginsDownloadURL.AddItem "http://www.computec.ch/projekte/atk/plugins/pluginslist/"
    
    'Add the search engine query urls
    cmbSearchEngineURL.AddItem "http://www.google.com/search?q="
    cmbSearchEngineURL.AddItem "http://search.yahoo.com/search?p="
    cmbSearchEngineURL.AddItem "http://www.hotbot.com/default.asp?query="
    cmbSearchEngineURL.AddItem "http://www.altavista.com/web/results?q="
    cmbSearchEngineURL.AddItem "http://www.alltheweb.com/search?q="
    cmbSearchEngineURL.AddItem "http://search.netscape.com/ns/search?query="
    cmbSearchEngineURL.AddItem "http://a9.com/"
    cmbSearchEngineURL.AddItem "http://search.msn.com/results.aspx?q="
    cmbSearchEngineURL.AddItem "http://msxml.excite.com/info.xcite/search/web/"
    cmbSearchEngineURL.AddItem "http://suche.fireball.de/cgi-bin/pursuit?query="
    cmbSearchEngineURL.AddItem "http://suche.lycos.de/cgi-bin/pursuit?query="
    cmbSearchEngineURL.AddItem "http://search.megaspider.com/XP.html?"
    cmbSearchEngineURL.AddItem "http://web.ask.com/web?q="
    cmbSearchEngineURL.AddItem "http://search.dmoz.org/cgi-bin/search?search="
    cmbSearchEngineURL.AddItem "http://astalavista.box.sk/cgi-bin/robot?srch="
    cmbSearchEngineURL.AddItem "http://www2.packetstormsecurity.org/cgi-bin/search/search.cgi?searchvalue="
    cmbSearchEngineURL.AddItem "http://astalavista.box.sk/cgi-bin/robot?srch="
    cmbSearchEngineURL.AddItem "http://search.gulli.com/"
    cmbSearchEngineURL.AddItem "http://www.gurunet.com/query?s="
    cmbSearchEngineURL.AddItem "http://froogle.google.com/froogle?q="
    cmbSearchEngineURL.AddItem "http://anon.free.anonymizer.com/http://www.google.com/search?q="

    cmbHelpURL.AddItem "http://www.computec.ch/projekte/atk/documentation/help/"

    'Show the configuration
    Call LoadActualConfigurationValues
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If CheckIfConfigIsEdited = True Then
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    'Check the window state. Do not resize if the window is minimized
    If Me.WindowState <> vbMinimized Then
        'Prevent zu small windows in height
        If Me.Height <> 4920 Then
            Me.Height = 4920
        End If
        
        'Prevent zu small windows in width
        If Me.Width <> 7215 Then
            Me.Width = 7215
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConfiguration = Nothing
End Sub

Private Sub lblOnlineHelpURLTest_Click()
    Dim strWebSiteURL As String
    
    strWebSiteURL = cmbHelpURL.Text
    
    'Load the online help
    WriteLogEntry "Loading the online help " & strWebSiteURL & " as a test ...", 6
    Call ShellExecute(Me.hwnd, "Open", strWebSiteURL, "", App.Path, 1)
End Sub

Private Sub lblReportTemplateNote_Click()
    'We are loading the form modal. This is not very nice but it is better than
    'another run-time error.
    frmReportConfiguration.Show vbModal
End Sub

Private Sub lblSearchEngineTest_Click()
    Dim strSearchEngineTestURL As String
    
    strSearchEngineTestURL = cmbSearchEngineURL.Text & "Attack Tool Kit"
    
    WriteLogEntry "Opening the search engine URL " & strSearchEngineTestURL & " for testing ...", 6
    Call ShellExecute(frmMain.hwnd, "Open", strSearchEngineTestURL, "", App.Path, 1)
End Sub

Private Sub lblTestSpeechFeature_Click()
    Dim bolActivateSpeechCache As Boolean
    
    If Not ActivateSpeech Then
        ActivateSpeech = True
    Else
        bolActivateSpeechCache = True
    End If
    
    Call ReadText("Test, check, check, one, two, ...")
    
    If bolActivateSpeechCache <> ActivateSpeech Then
        ActivateSpeech = bolActivateSpeechCache
    End If
End Sub

Private Sub mnuFileCloseItem_Click()
    Unload Me
End Sub

Private Sub mnuFileNewItem_Click()
    Call LoadConfigFromFile(App.Path & "\configs\127.0.0.1-" & Date & ".config")

    'Show the "new" configuration
    Call LoadActualConfigurationValues
End Sub

Private Sub mnuFileOpenItem_Click()
    Dim strConfigurationFileName As String  'The name of the configuration file
    Dim strDefaultConfigurationPath As String
    
    strDefaultConfigurationPath = App.Path & "\configs"
    
    'Define the initial directory of the plugins
    On Error Resume Next
    If Not (Dir$(strDefaultConfigurationPath, 16) <> "") Then
        strDefaultConfigurationPath = App.Path
    End If
    
    cdgOpen.Filename = application_configuration_filename
    
    'Define the initial directory of the plugins
    cdgOpen.InitDir = strDefaultConfigurationPath
    
    'Ask the user for the desired filename
    cdgOpen.ShowOpen 'Opens the save dialog
    
    'Cache the filename into a variant to increase the speed
    strConfigurationFileName = cdgOpen.Filename
    
    'Check if a file was selected
    If LenB(strConfigurationFileName) Then
        'Check if the file exists
        If (Dir$(strConfigurationFileName, 16) <> "") Then
            'Load the configuration file
            Call LoadConfigFromFile(strConfigurationFileName)
            Call LoadActualConfigurationValues
        End If
    End If
End Sub

Private Sub mnuFileSaveAsItem_Click()
    Dim strConfigurationFileName As String    'Here we save the desired filename for the new plugin
    Dim strDefaultConfigurationPath As String
    
    strDefaultConfigurationPath = App.Path & "\configs"
    
    'Define the initial directory of the plugins
    On Error Resume Next
    If Not (Dir$(strDefaultConfigurationPath, 16) <> "") Then
        strDefaultConfigurationPath = App.Path
    End If
    
    cdgSaveAs.InitDir = strDefaultConfigurationPath
    
    'Ask the user for the desired filename
    cdgSaveAs.Filename = Target & "-" & Date & ".config"
    On Error GoTo ErrSub
    cdgSaveAs.ShowSave 'Opens the save dialog
    strConfigurationFileName = cdgSaveAs.Filename 'Get the filename
    
    'Cut the plugin extension if there is one given
    If LenB(strConfigurationFileName) Then
        Call SaveConfigurationData
        Call WriteConfigurationToFile(strConfigurationFileName)
    End If
ErrSub:
End Sub

Private Sub mnuFileSaveItem_Click()
    Call SaveConfigurationData
    Call WriteConfigurationToFile(application_configuration_filename)
End Sub

Private Sub mnuHelpConfigurationHelpItem_Click()
    Call OpenOnlineHelp("configuration")
End Sub

Private Sub optFullAudit_Click()
    MsgBox "The full audit mode is not the main feature of the ATK." & vbCrLf & _
        "The mode is not very efficient and a general audit task" & vbCrLf & _
        "can much better be done by other well-known security scanners." & vbCrLf & _
        "Please use the single check mode for checking dedicated" & vbCrLf & _
        "vulnerabilities (perhaps already identified by other scanners)" & vbCrLf & _
        "instead.", _
        vbInformation, "Attack Tool Kit full audit information"
End Sub

Private Sub tspConfiguration_Click()
    'Target
    If tspConfiguration.SelectedItem.Index = 1 Then
        fraTarget.Visible = True
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Preferences
    ElseIf tspConfiguration.SelectedItem.Index = 2 Then
        fraPreferences.Visible = True
        fraTarget.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Mapping
    ElseIf tspConfiguration.SelectedItem.Index = 3 Then
        fraMapping.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Plugins
    ElseIf tspConfiguration.SelectedItem.Index = 4 Then
        fraPlugins.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Alerting
    ElseIf tspConfiguration.SelectedItem.Index = 5 Then
        fraAlerting.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraSuggestions.Visible = True
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Suggestions
    ElseIf tspConfiguration.SelectedItem.Index = 6 Then
        fraSuggestions.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False

    'Searchengine
    ElseIf tspConfiguration.SelectedItem.Index = 7 Then
        fraSearchengine.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraReports.Visible = False
        fraSuggestions.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Reports
    ElseIf tspConfiguration.SelectedItem.Index = 8 Then
        fraReports.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Logs
    ElseIf tspConfiguration.SelectedItem.Index = 9 Then
        fraLogs.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Speech
    ElseIf tspConfiguration.SelectedItem.Index = 10 Then
        fraSpeech.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraHelp.Visible = False
    
    'Help
    ElseIf tspConfiguration.SelectedItem.Index = 11 Then
        fraHelp.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraSpeech.Visible = False
        fraLogs.Visible = False
    
    End If
End Sub

Private Sub txtDefaultSleep_Change()
    If Len(txtDefaultSleep.Text) < 1 Then
        txtDefaultSleep.Text = 1
    Else
        If txtDefaultSleep.Text < 1 Then
            txtDefaultSleep.Text = 1
        End If
    End If
End Sub

Private Sub txtDefaultSleep_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyBack
    Case Else
        KeyAscii = 0
  End Select
End Sub

Private Sub txtTarget_LostFocus()
    Dim strNewTarget As String
    
    strNewTarget = txtTarget.Text
    
    If Mid$(strNewTarget, 1, 7) = "http://" Then
        strNewTarget = Mid$(strNewTarget, 8, Len(strNewTarget))
        Call errTargetWrongSpecification
    ElseIf Mid$(strNewTarget, 1, 6) = "ftp://" Then
        strNewTarget = Mid$(strNewTarget, 7, Len(strNewTarget))
        Call errTargetWrongSpecification
    ElseIf Mid$(strNewTarget, 1, 2) = "\\" Then
        strNewTarget = Mid$(strNewTarget, 3, Len(strNewTarget))
        Call errTargetWrongSpecification
    End If
    
    txtTarget.Text = strNewTarget
    
    If Len(strNewTarget) < 1 Then
        MsgBox ("Target missing." & vbCrLf & vbCrLf & _
            "Please enter the host name or IP address of the target."), vbInformation, "Attack Tool Kit error"
        txtTarget.SetFocus
    End If
End Sub

Private Sub txtTimeout_Change()
    If Len(txtTimeout.Text) < 1 Then
        txtTimeout.Text = 10000
    Else
        If txtTimeout.Text < 10000 Then
            txtTimeout.Text = 10000
        End If
    End If
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function CheckIfConfigIsEdited() As Boolean
    Dim iMsgBoxResponse As Integer
    
    CheckIfConfigIsEdited = False
    
    iMsgBoxResponse = MsgBox("You have changed the behavior of the software by" & vbCrLf & _
            "changing the configuration." & vbCrLf & vbCrLf & _
            "Would you like to save the existing configuration?", _
            vbYesNoCancel + vbInformation, "Attack Tool Kit configuration changed")
            
    If iMsgBoxResponse = vbYes Then
        Call mnuFileSaveItem_Click
        Unload Me
    ElseIf iMsgBoxResponse = vbNo Then
        Unload Me
    ElseIf iMsgBoxResponse = vbCancel Then
        CheckIfConfigIsEdited = True
    End If
End Function

Private Sub SaveConfigurationData()
    'Write the new values
    Target = txtTarget.Text
    
    AttackTimeout = txtTimeout.Text
    DefaultSleepValue = txtDefaultSleep.Text

    If PluginDirectory <> dirPlugins.Path Then
        PluginDirectory = dirPlugins.Path
        frmMain.filATKPlugins.Path = PluginDirectory
        'Call frmMain.mnuPluginsReloadAllItem_Click
    End If
    
    If PluginsDownloadURL <> cmbPluginsDownloadURL.Text Then
        PluginsDownloadURL = cmbPluginsDownloadURL.Text
    End If
    
    If SearchEngineURL <> cmbSearchEngineURL.Text Then
        SearchEngineURL = cmbSearchEngineURL.Text
    End If
    
    If HelpURL <> cmbHelpURL.Text Then
        HelpURL = cmbHelpURL.Text
    End If
    
    If SuggestionsDirectory <> dirSuggestions.Path Then
        SuggestionsDirectory = dirSuggestions.Path
    End If

    'If ReportsDirectory <> dirReportsDirectory.Path Then
    '    ReportsDirectory = dirReportsDirectory.Path
    'End If

    If optSingleCheck.Value = True Then
        AttackMode = "SingleCheck"
    ElseIf optFullAudit.Value = True Then
        AttackMode = "FullAudit"
    End If

    If chkDoSilentChecks.Value = 1 Then
        DoSilentChecks = True
    Else
        DoSilentChecks = False
    End If
    
    If chkDoNoDoSChecks.Value = 1 Then
        DoNoDoSChecks = True
    Else
        DoNoDoSChecks = False
    End If

    If chkDoICMPMapping.Value = 1 Then
        DoICMPMapping = True
    Else
        DoICMPMapping = False
    End If

    If chkScanifICMPfails.Value = 1 Then
        ScanIfICMPFails = True
    Else
        ScanIfICMPFails = False
    End If
    
    If chkSuggestions.Value = 1 Then
        ActivateSuggestions = True
    Else
        ActivateSuggestions = False
    End If
        
    If chkAlertingVulnerabilityFound.Value = 1 Then
        AlertingVulnFound = True
    Else
        AlertingVulnFound = False
    End If
    
    If chkAlertingVulnerabilityNotFound.Value = 1 Then
        AlertingVulnNotFound = True
    Else
        AlertingVulnNotFound = False
    End If
    
    If chkActivateLogs.Value = 1 Then
        ActivateLogs = True
    Else
        ActivateLogs = False
    End If
    If InStr(1, cmbLogsSecurityLevel.Text, "0 ", vbBinaryCompare) Then
        LogsSecurityLevel = 0
    ElseIf InStr(1, cmbLogsSecurityLevel.Text, "1 ", vbBinaryCompare) Then
        LogsSecurityLevel = 1
    ElseIf InStr(1, cmbLogsSecurityLevel.Text, "2 ", vbBinaryCompare) Then
        LogsSecurityLevel = 2
    ElseIf InStr(1, cmbLogsSecurityLevel.Text, "3 ", vbBinaryCompare) Then
        LogsSecurityLevel = 3
    ElseIf InStr(1, cmbLogsSecurityLevel.Text, "4 ", vbBinaryCompare) Then
        LogsSecurityLevel = 4
    ElseIf InStr(1, cmbLogsSecurityLevel.Text, "5 ", vbBinaryCompare) Then
        LogsSecurityLevel = 5
    ElseIf InStr(1, cmbLogsSecurityLevel.Text, "6 ", vbBinaryCompare) Then
        LogsSecurityLevel = 6
    Else
        LogsSecurityLevel = 7
    End If
    LogsDirectory = dirLogs.Path
    
    If chkActivateSpeech.Value = 1 Then
        ActivateSpeech = True
    Else
        ActivateSpeech = False
    End If

    If IsFormVisible("frmAttackVisualizing") = True Then
        frmAttackVisualizing.txtTargetData.Text = Target
        
        If InStr(1, Target, "192.") <> 0 Then
            frmAttackVisualizing.lblNetworkName.Caption = "LAN"
        ElseIf InStr(1, Target, "172.") <> 0 Then
            frmAttackVisualizing.lblNetworkName.Caption = "LAN"
        ElseIf InStr(1, Target, "10.") <> 0 Then
            frmAttackVisualizing.lblNetworkName.Caption = "LAN"
        ElseIf InStr(1, Target, "127.") <> 0 Then
            frmAttackVisualizing.lblNetworkName.Caption = "Localhost"
        Else
            frmAttackVisualizing.lblNetworkName.Caption = "Internet"
        End If
        
    End If
End Sub
