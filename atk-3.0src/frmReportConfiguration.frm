VERSION 5.00
Begin VB.Form frmReportConfiguration 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report Configuration"
   ClientHeight    =   7455
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Top             =   6960
      Width           =   855
   End
   Begin VB.Frame fraHeaderCustomizing 
      Caption         =   "Header Customizing"
      Height          =   2775
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   6135
      Begin VB.CommandButton cmdHeaderAdd 
         Caption         =   "+"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdHeaderRemove 
         Caption         =   "-"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   1680
         Width           =   375
      End
      Begin VB.ListBox lstHeaderReport 
         Height          =   2010
         Left            =   3120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   2415
      End
      Begin VB.ListBox lstHeaderPositions 
         Height          =   2010
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdHeaderUp 
         Caption         =   "up"
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdHeaderDown 
         Caption         =   "dn"
         Height          =   255
         Left            =   5640
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Available header positions:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Actual header report structure:"
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   9120
      TabIndex        =   16
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   10080
      TabIndex        =   17
      Top             =   6960
      Width           =   855
   End
   Begin VB.Frame fraExample 
      Caption         =   "Example"
      Height          =   6735
      Left            =   6360
      TabIndex        =   19
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtReportExample 
         Height          =   6375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame fraVulnerabilitiesCustomizing 
      Caption         =   "Vulnerabilities Customizing"
      Height          =   2775
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   6135
      Begin VB.CommandButton cmdVulnerabilityDown 
         Caption         =   "dn"
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdVulnerabilityUp 
         Caption         =   "up"
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   1320
         Width           =   375
      End
      Begin VB.ListBox lstVulnerabilityPositions 
         Height          =   2010
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   2415
      End
      Begin VB.ListBox lstVulnerabilityReport 
         Height          =   2010
         Left            =   3120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdVulnerabilityRemove 
         Caption         =   "-"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdVulnerabilityAdd 
         Caption         =   "+"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Actual vulnerability report structure:"
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Available vulnerability positions:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame fraTemplates 
      Caption         =   "Templates"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox cmbReportTemplates 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblTemplateDescription 
         Caption         =   "Please select a report template or create your own report."
         Height          =   495
         Left            =   2760
         TabIndex        =   26
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Label lblDragAndDrop 
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCloseItem 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpReportConfigurationHelpItem 
         Caption         =   "Report Configuration Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmReportConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Frame Description                                                                *
' *                                                                                  *
' * In this frame the user is able to configure the whole reporting.                 *
' ************************************************************************************

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 3.0 2004-11-01                                                           *
' * - Replaced all useless functions with normal subs.                               *
' * - Corrected the tab order to be a more logical.                                  *
' * Version 2.0 2004-08-15                                                           *
' * - Added the whole new fields and some special fields fur further diagnostics.    *
' * - Fixed a nasty error when the up and down buttons were pushed too much.         *
' * - Added and corrected the whole tab stops.                                       *
' ************************************************************************************

Private Sub cmbReportTemplates_Click()
    'Developer Note: If somebody knows the report structure of other, not listed
    'vulnerability scanners, please send me such a report. So I could the
    'missing one too.
    
    'Load the wanted templates
    If cmbReportTemplates.Text = "ATK technical report" Then
        lblTemplateDescription.Caption = "The ATK report for technical employees. Good for further analysis."
        Call LoadATKTechnicalReport
    ElseIf cmbReportTemplates.Text = "ATK bugfix report" Then
        lblTemplateDescription.Caption = "The ATK report for technical bugfixing. Good for fixing the potential flaws."
        Call LoadATKBugfixingReport
    ElseIf cmbReportTemplates.Text = "ATK management report" Then
        lblTemplateDescription.Caption = "The ATK report for the management. Good to get a quick overview. No technical details."
        Call LoadATKManagementReport
    ElseIf cmbReportTemplates.Text = "Nessus old report" Then
        lblTemplateDescription.Caption = "The old Nessus report. Good for further analysis."
        Call LoadNessusOldReport
    ElseIf cmbReportTemplates.Text = "Nessus new report" Then
        lblTemplateDescription.Caption = "The new Nessus report. Good for further analysis."
        Call LoadNessusNewReport
    ElseIf cmbReportTemplates.Text = "Symantec NetRecon report" Then
        lblTemplateDescription.Caption = "The Symantec NetRecon report. Good for further analysis and fixing."
        Call LoadSymantecNetreconReport
    End If
End Sub

Private Sub cmbReportTemplates_KeyPress(KeyAscii As Integer)
    'Do not allow manual input
    KeyAscii = 0
End Sub

Private Sub cmdAccept_Click()
    'Recompute the report structure
    Call PrepareReportStructure

    'Show the new report structure also in the main frome
    frmMain.txtPluginContent.Text = PluginReportData

    'Close the window
    Call cmdClose_Click
End Sub

Private Sub cmdHeaderAdd_Click()
    lstHeaderReport.AddItem lstHeaderPositions.Text
    'After changing the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdVulnerabilityAdd_Click()
    lstVulnerabilityReport.AddItem lstVulnerabilityPositions.Text
    
    'After changing the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub cmdVulnerabilityDown_Click()
    Dim strTemp1 As String    'Hold the selected index data temporarily for move
    Dim iCnt    As Integer    'Holds the index of the item to be moved
        
    'Assign the first index
    iCnt = lstVulnerabilityReport.ListIndex
    
    If iCnt < lstVulnerabilityReport.ListCount - 1 Then
         
        strTemp1 = lstVulnerabilityReport.List(iCnt)
        
        'Add the item selected to below the current position
        lstVulnerabilityReport.AddItem strTemp1, (iCnt + 2)
        
        lstVulnerabilityReport.RemoveItem (iCnt)
        
        'Reselect the item that was moved.
        lstVulnerabilityReport.Selected(iCnt + 1) = True
    End If

    'After chaging the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub cmdHeaderDown_Click()
    Dim strTemp1 As String    'Hold the selected index data temporarily for move
    Dim iCnt    As Integer    'Holds the index of the item to be moved
        
    'Assign the first index
    iCnt = lstHeaderReport.ListIndex
    
    If iCnt < lstHeaderReport.ListCount - 1 Then
        strTemp1 = lstHeaderReport.List(iCnt)
        
        'Add the item selected to below the current position
        lstHeaderReport.AddItem strTemp1, (iCnt + 2)
        
        lstHeaderReport.RemoveItem (iCnt)
        
        'Reselect the item that was moved.
        lstHeaderReport.Selected(iCnt + 1) = True
   End If

    'After chaging the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub cmdRefresh_Click()
    'Clear the example report
    txtReportExample.Text = vbNullString

    txtReportExample.Text = PluginReportData
End Sub

Private Sub cmdVulnerabilityRemove_Click()
    'Prevent errors if no item is selected
    If lstVulnerabilityReport.ListIndex < 0 Then
        lstVulnerabilityReport.ListIndex = lstVulnerabilityReport.ListCount - 1
    End If
        
    'Delete the selected item
    If lstVulnerabilityReport.ListCount > 0 Then
        lstVulnerabilityReport.RemoveItem lstVulnerabilityReport.ListIndex
    End If

    'After chaging the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub cmdHeaderRemove_Click()
    'Prevent errors if no item is selected
    If lstHeaderReport.ListIndex < 0 Then
        lstHeaderReport.ListIndex = lstHeaderReport.ListCount - 1
    End If
        
    'Delete the selected item
    If lstHeaderReport.ListCount > 0 Then
        lstHeaderReport.RemoveItem lstHeaderReport.ListIndex
    End If

    'After chaging the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub cmdVulnerabilityUp_Click()
    Dim strTemp1 As String   '-- hold the selected index data temporarily for move
    Dim iCnt    As Integer   '-- holds the index of the item to be moved
    iCnt = lstVulnerabilityReport.ListIndex
        
        If iCnt > 0 Then
         
        strTemp1 = lstVulnerabilityReport.List(iCnt)
        
        'Add the item selected to one position above the current position
        lstVulnerabilityReport.AddItem strTemp1, (iCnt - 1)
        
        'remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstVulnerabilityReport.RemoveItem (iCnt + 1)
        
        'Reselect the item that was moved.
        lstVulnerabilityReport.Selected(iCnt - 1) = True
    End If

    'After chaging the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub cmdHeaderUp_Click()
    Dim strTemp1 As String   '-- hold the selected index data temporarily for move
    Dim iCnt    As Integer   '-- holds the index of the item to be moved
    iCnt = lstHeaderReport.ListIndex
    
    
    If iCnt > 0 Then
         
        strTemp1 = lstHeaderReport.List(iCnt)
        
        'Add the item selected to one position above the current position
        lstHeaderReport.AddItem strTemp1, (iCnt - 1)
        
        'remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstHeaderReport.RemoveItem (iCnt + 1)
        
        'Reselect the item that was moved.
        lstHeaderReport.Selected(iCnt - 1) = True
    End If

    'After chaging the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If State = 0 Then Source.MousePointer = 12
    If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub Form_Load()
    'Add the report templates in the combobox
    cmbReportTemplates.AddItem "ATK technical report"
    cmbReportTemplates.AddItem "ATK management report"
    cmbReportTemplates.AddItem "ATK bugfix report"
'    cmbReportTemplates.AddItem "ISS Internet Scanner report"
    cmbReportTemplates.AddItem "Nessus new report"
    cmbReportTemplates.AddItem "Nessus old report"
    cmbReportTemplates.AddItem "Symantec NetRecon report"
    
    'Pre-select the first entry
    cmbReportTemplates.ListIndex = 0
    
    'Add the items for the header positions
    lstHeaderPositions.AddItem "<br>"
    lstHeaderPositions.AddItem "scan_target"
    lstHeaderPositions.AddItem "scan_time"
    lstHeaderPositions.AddItem "scan_date"
    lstHeaderPositions.AddItem "scan_mode"
    lstHeaderPositions.AddItem "software_version"
    lstHeaderPositions.AddItem "software_pluginscount"

    'Add the items for the vulnerability positions
    lstVulnerabilityPositions.AddItem "<br>"
    
    lstVulnerabilityPositions.AddItem "plugin_id"
    lstVulnerabilityPositions.AddItem "plugin_name"
    lstVulnerabilityPositions.AddItem "plugin_filename"
    lstVulnerabilityPositions.AddItem "plugin_filesize"
    lstVulnerabilityPositions.AddItem "plugin_family"
    lstVulnerabilityPositions.AddItem "plugin_created_name"
    lstVulnerabilityPositions.AddItem "plugin_created_email"
    lstVulnerabilityPositions.AddItem "plugin_created_web"
    lstVulnerabilityPositions.AddItem "plugin_created_company"
    lstVulnerabilityPositions.AddItem "plugin_created_date"
    lstVulnerabilityPositions.AddItem "plugin_updated_name"
    lstVulnerabilityPositions.AddItem "plugin_updated_email"
    lstVulnerabilityPositions.AddItem "plugin_updated_web"
    lstVulnerabilityPositions.AddItem "plugin_updated_company"
    lstVulnerabilityPositions.AddItem "plugin_updated_date"
    lstVulnerabilityPositions.AddItem "plugin_version"
    lstVulnerabilityPositions.AddItem "plugin_changelog"
    lstVulnerabilityPositions.AddItem "plugin_protocol"
    lstVulnerabilityPositions.AddItem "plugin_port"
    lstVulnerabilityPositions.AddItem "plugin_procedure_detection"
    lstVulnerabilityPositions.AddItem "plugin_procedure_exploit"
    lstVulnerabilityPositions.AddItem "plugin_detection_accuracy"
    lstVulnerabilityPositions.AddItem "plugin_exploit_accuracy"
    lstVulnerabilityPositions.AddItem "plugin_comment"
    
    lstVulnerabilityPositions.AddItem "bug_published_name"
    lstVulnerabilityPositions.AddItem "bug_published_email"
    lstVulnerabilityPositions.AddItem "bug_published_web"
    lstVulnerabilityPositions.AddItem "bug_published_company"
    lstVulnerabilityPositions.AddItem "bug_published_date"
    lstVulnerabilityPositions.AddItem "bug_advisory"
    lstVulnerabilityPositions.AddItem "bug_produced_name"
    lstVulnerabilityPositions.AddItem "bug_produced_email"
    lstVulnerabilityPositions.AddItem "bug_produced_web"
    lstVulnerabilityPositions.AddItem "bug_affected"
    lstVulnerabilityPositions.AddItem "bug_not_affected"
    lstVulnerabilityPositions.AddItem "bug_false_positives"
    lstVulnerabilityPositions.AddItem "bug_false_negatives"
    lstVulnerabilityPositions.AddItem "bug_vulnerability_class"
    lstVulnerabilityPositions.AddItem "bug_local"
    lstVulnerabilityPositions.AddItem "bug_remote"
    lstVulnerabilityPositions.AddItem "bug_description"
    lstVulnerabilityPositions.AddItem "bug_response"
    lstVulnerabilityPositions.AddItem "bug_solution"
    lstVulnerabilityPositions.AddItem "bug_fixing_time"
    lstVulnerabilityPositions.AddItem "bug_exploit_availability"
    lstVulnerabilityPositions.AddItem "bug_exploit_url"
    lstVulnerabilityPositions.AddItem "bug_severity"
    lstVulnerabilityPositions.AddItem "bug_popularity"
    lstVulnerabilityPositions.AddItem "bug_simplicity"
    lstVulnerabilityPositions.AddItem "bug_impact"
    lstVulnerabilityPositions.AddItem "bug_risk"
    lstVulnerabilityPositions.AddItem "bug_nessus_risk"
    lstVulnerabilityPositions.AddItem "bug_iss_scanner_rating"
    lstVulnerabilityPositions.AddItem "bug_netrecon_rating"
    lstVulnerabilityPositions.AddItem "bug_checking_tool"

    lstVulnerabilityPositions.AddItem "source_cve"
    lstVulnerabilityPositions.AddItem "source_certvu_id"
    lstVulnerabilityPositions.AddItem "source_cert_id"
    lstVulnerabilityPositions.AddItem "source_uscertta_id"
    lstVulnerabilityPositions.AddItem "source_securityfocus_bid"
    lstVulnerabilityPositions.AddItem "source_osvdb_id"
    lstVulnerabilityPositions.AddItem "source_secunia_id"
    lstVulnerabilityPositions.AddItem "source_securiteam_url"
    lstVulnerabilityPositions.AddItem "source_securitytracker_id"
    lstVulnerabilityPositions.AddItem "source_scip_id"
    lstVulnerabilityPositions.AddItem "source_tecchannel_id"
    lstVulnerabilityPositions.AddItem "source_heise_news"
    lstVulnerabilityPositions.AddItem "source_heise_security"
    lstVulnerabilityPositions.AddItem "source_aerasec_id"
    lstVulnerabilityPositions.AddItem "source_nessus_id"
    lstVulnerabilityPositions.AddItem "source_issxforce_id"
    lstVulnerabilityPositions.AddItem "source_snort_id"
    lstVulnerabilityPositions.AddItem "source_arachnids_id"
    lstVulnerabilityPositions.AddItem "source_mssb_id"
    lstVulnerabilityPositions.AddItem "source_mskb_id"
    lstVulnerabilityPositions.AddItem "source_netbsdsa_id"
    lstVulnerabilityPositions.AddItem "source_rhsa_id"
    lstVulnerabilityPositions.AddItem "source_ciac_id"
    lstVulnerabilityPositions.AddItem "source_literature"
    lstVulnerabilityPositions.AddItem "source_misc"
    
    lstVulnerabilityPositions.AddItem "session_procedure_type"
    lstVulnerabilityPositions.AddItem "session_procedure_commands"
    
    lstVulnerabilityPositions.AddItem "report_header_structure"
    lstVulnerabilityPositions.AddItem "report_vulnerability_structure"
    
    'Select the first item in the positions
    lstHeaderPositions.ListIndex = 0
    lstVulnerabilityPositions.ListIndex = 0
    
    'Refresh the example report
    Call cmdRefresh_Click
End Sub

Private Sub lstHeaderPositions_Click()
   'Clear the example report
    txtReportExample.Text = vbNullString
        
    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub lstVulnerabilityPositions_Click()
    'Clear the example report
    txtReportExample.Text = vbNullString
    
    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub lstVulnerabilityPositions_DblClick()
    lstVulnerabilityReport.AddItem lstVulnerabilityPositions.Text

    'After chaging the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub lstHeaderPositions_DblClick()
    lstHeaderReport.AddItem lstHeaderPositions.Text
    
    'After chaging the template do rename the report name
    Call LoadATKCustomReport
End Sub

Private Sub lstVulnerabilityPositions_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If State = 0 Then Source.MousePointer = 12
    If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub lstHeaderPositions_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If State = 0 Then Source.MousePointer = 12
    If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub lstHeaderPositions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Dim DY As Integer
        
        DY = TextHeight("A")
        lblDragAndDrop.Move fraHeaderCustomizing.Left + lstHeaderPositions.Left, fraHeaderCustomizing.Top + lstHeaderPositions.Top + y - DY * 0.5, lstHeaderPositions.Width, DY
        lblDragAndDrop.Drag
        
        Call LoadATKCustomReport
    End If
End Sub

Private Sub LoadATKCustomReport()
    'After chaging the template do rename the report name
    cmbReportTemplates.Text = "ATK custom report"
    
    lblTemplateDescription.Caption = "No template loaded. You are creating a custom report."
End Sub

Private Sub lstVulnerabilityPositions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Dim DY As Integer
        
        DY = TextHeight("A")
        lblDragAndDrop.Move fraVulnerabilitiesCustomizing.Left + lstVulnerabilityPositions.Left, _
            fraVulnerabilitiesCustomizing.Top + lstVulnerabilityPositions.Top + y - DY * 0.5, lstVulnerabilityPositions.Width, DY
        lblDragAndDrop.Drag
        
        'After chaging the template do rename the report name
        Call LoadATKCustomReport
    End If
End Sub

Private Sub lstVulnerabilityReport_Click()
    'Clear the example report
    txtReportExample.Text = vbNullString
    
    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub lstHeaderReport_Click()
    'Clear the example report
    txtReportExample.Text = vbNullString
    
    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub lstVulnerabilityReport_DblClick()
    lstVulnerabilityReport.RemoveItem lstVulnerabilityReport.ListIndex
    
    'Clear the example report
    txtReportExample.Text = vbNullString
    
    'After chaging the template do rename the report name
    Call LoadATKCustomReport
    
    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub lstHeaderReport_DblClick()
    lstHeaderReport.RemoveItem lstHeaderReport.ListIndex
    
    'Clear the example report
    txtReportExample.Text = vbNullString
    
    'After chaging the template do rename the report name
    Call LoadATKCustomReport
    
    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub lstVulnerabilityReport_DragDrop(Source As Control, x As Single, y As Single)
    lstVulnerabilityReport.AddItem lstVulnerabilityPositions.Text
End Sub

Private Sub lstHeaderReport_DragDrop(Source As Control, x As Single, y As Single)
    lstHeaderReport.AddItem lstHeaderPositions.Text
End Sub

Private Sub LoadATKTechnicalReport()
    'Clear the last input
    lstHeaderReport.Clear
    lstVulnerabilityReport.Clear
    
    'Add the new header input for technicians
    lstHeaderReport.AddItem "scan_target"
    lstHeaderReport.AddItem "scan_time"
    lstHeaderReport.AddItem "scan_date"
    lstHeaderReport.AddItem "scan_mode"
    lstHeaderReport.AddItem "<br>"
    lstHeaderReport.AddItem "software_version"
    lstHeaderReport.AddItem "software_pluginscount"
    lstHeaderReport.AddItem "<br>"
    lstHeaderReport.AddItem "<br>"
    
    'Add the new vulnerability input for technicians
    lstVulnerabilityReport.AddItem "plugin_id"
    lstVulnerabilityReport.AddItem "plugin_name"
    lstVulnerabilityReport.AddItem "plugin_protocol"
    lstVulnerabilityReport.AddItem "plugin_port"
    lstVulnerabilityReport.AddItem "bug_severity"
    lstVulnerabilityReport.AddItem "bug_advisory"
    lstVulnerabilityReport.AddItem "bug_affected"
    lstVulnerabilityReport.AddItem "bug_not_affected"
    lstVulnerabilityReport.AddItem "bug_vulnerability_class"
    lstVulnerabilityReport.AddItem "bug_exploit_url"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_description"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_response"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_solution"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "source_cve"
    lstVulnerabilityReport.AddItem "source_securityfocus_bid"
    lstVulnerabilityReport.AddItem "source_nessus_id"
    
    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub LoadATKBugfixingReport()
    'Clear the last input
    lstHeaderReport.Clear
    lstVulnerabilityReport.Clear

    'Add the new header input for bugfixing
    lstHeaderReport.AddItem "scan_target"
    lstHeaderReport.AddItem "scan_time"
    lstHeaderReport.AddItem "scan_date"
    lstHeaderReport.AddItem "scan_mode"
    lstHeaderReport.AddItem "<br>"
    lstHeaderReport.AddItem "software_version"
    lstHeaderReport.AddItem "software_pluginscount"
    lstHeaderReport.AddItem "<br>"
    lstHeaderReport.AddItem "<br>"
    
    'Add the vulnerability information for bugfixing
    lstVulnerabilityReport.AddItem "plugin_id"
    lstVulnerabilityReport.AddItem "plugin_name"
    lstVulnerabilityReport.AddItem "plugin_protocol"
    lstVulnerabilityReport.AddItem "plugin_port"
    lstVulnerabilityReport.AddItem "bug_severity"
    lstVulnerabilityReport.AddItem "bug_affected"
    lstVulnerabilityReport.AddItem "bug_not_affected"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_description"
    lstVulnerabilityReport.AddItem "bug_check_tool"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_solution"
    lstVulnerabilityReport.AddItem "bug_fixing_time"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "source_cve"
    lstVulnerabilityReport.AddItem "source_nessus_id"
    lstVulnerabilityReport.AddItem "snort_id"
    lstVulnerabilityReport.AddItem "source_literature"
    lstVulnerabilityReport.AddItem "source_misc"

    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub LoadATKManagementReport()
    'Clear the last input
    lstHeaderReport.Clear
    lstVulnerabilityReport.Clear
    
    'Add the new header input for the management
    lstHeaderReport.AddItem "scan_target"
    lstHeaderReport.AddItem "scan_time"
    lstHeaderReport.AddItem "scan_date"
    lstHeaderReport.AddItem "<br>"
    lstHeaderReport.AddItem "software_version"
    lstHeaderReport.AddItem "<br>"
    lstHeaderReport.AddItem "<br>"
    
    'Add the new vulnerability input for the management
    lstVulnerabilityReport.AddItem "plugin_name"
    lstVulnerabilityReport.AddItem "bug_severity"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_description"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_solution"

    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub LoadNessusOldReport()
    'Clear the last input
    lstHeaderReport.Clear
    lstVulnerabilityReport.Clear

    'Add the new header input a la Nessus
    lstHeaderReport.AddItem "scan_target"
    lstHeaderReport.AddItem "<br>"
    
    'Add the new vulnerability input a la Nessus
    lstVulnerabilityReport.AddItem "plugin_name"
    lstVulnerabilityReport.AddItem "plugin_port"
    lstVulnerabilityReport.AddItem "plugin_protocol"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_description"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_solution"
    lstVulnerabilityReport.AddItem "bug_severity"
    lstVulnerabilityReport.AddItem "source_cve"
    lstVulnerabilityReport.AddItem "source_securityfocus_bid"
    lstVulnerabilityReport.AddItem "source_nessus_id"

    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub LoadNessusNewReport()
    'Clear the last input
    lstHeaderReport.Clear
    lstVulnerabilityReport.Clear

    'Add the new header input a la Nessus
    lstHeaderReport.AddItem "software_version"
    lstHeaderReport.AddItem "<br>"
    lstHeaderReport.AddItem "scan_target"
    lstHeaderReport.AddItem "scan_mode"
    lstHeaderReport.AddItem "<br>"
    lstHeaderReport.AddItem "<br>"
    
    'Add the new vulnerability input a la Nessus
    lstVulnerabilityReport.AddItem "bug_vulnerability_class"
    lstVulnerabilityReport.AddItem "plugin_port"
    lstVulnerabilityReport.AddItem "plugin_protocol"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_description"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_severity"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_solution"
    lstVulnerabilityReport.AddItem "source_cve"
    lstVulnerabilityReport.AddItem "source_securityfocus_bid"
    lstVulnerabilityReport.AddItem "source_nessus_id"

    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub LoadSymantecNetreconReport()
    'Clear the last input
    lstHeaderReport.Clear
    lstVulnerabilityReport.Clear
    
    'Add the new header input a la Symantec NetRecon 3.x
    lstHeaderReport.AddItem "scan_target"
    lstHeaderReport.AddItem "software_version"
    lstHeaderReport.AddItem "scan_time"
    lstHeaderReport.AddItem "scan_date"
    lstHeaderReport.AddItem "scan_mode"
    lstHeaderReport.AddItem "software_pluginscount"
    
    'Add the new vulnerability a la Symantec NetRecon 3.x
    lstVulnerabilityReport.AddItem "bug_affected"
    lstHeaderReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "plugin_name"
    lstVulnerabilityReport.AddItem "bug_severity"
    lstVulnerabilityReport.AddItem "bug_description"
    lstVulnerabilityReport.AddItem "bug_solution"
    lstVulnerabilityReport.AddItem "source_cve"
    lstVulnerabilityReport.AddItem "source_securityfocus_bid"
    lstVulnerabilityReport.AddItem "source_secunia_id"
    lstVulnerabilityReport.AddItem "source_scip_id"
    lstVulnerabilityReport.AddItem "source_tecchannel_id"
    lstVulnerabilityReport.AddItem "source_heise_news"
    lstVulnerabilityReport.AddItem "source_heise_security"
    lstVulnerabilityReport.AddItem "source_aerasec_id"
    lstVulnerabilityReport.AddItem "<br>"
    lstVulnerabilityReport.AddItem "bug_response"
        
    'Refresh the example report
    cmdRefresh_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLogEntry "Unloading the " & Me.Caption & " window.", 6
    Set frmReportConfiguration = Nothing
End Sub

Private Sub mnuFileCloseItem_Click()
    Call cmdAccept_Click
End Sub

Private Sub mnuHelpReportConfigurationHelpItem_Click()
    Call OpenOnlineHelp("report_configuration")
End Sub
