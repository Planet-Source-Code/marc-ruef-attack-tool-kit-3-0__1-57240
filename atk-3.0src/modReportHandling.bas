Attribute VB_Name = "modReportHandling"
Option Explicit

' ************************************************************************************
' * Frame Description                                                                *
' *                                                                                  *
' * In this frame the user is able to configure the report structure.                *
' ************************************************************************************

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 2.0 2004-08-24                                                           *
' * - A nasty bug was fixed. The last entry was not computed and missing.            *
' ************************************************************************************

Public report_header_structure As String
Public report_vulnerability_structure As String
Public ActualReport As String

Public Function LoadDefaultReportStructure()
    report_header_structure = _
    "H=scan_target" & vbCrLf & _
    "H=scan_time" & vbCrLf & _
    "H=scan_date" & vbCrLf & _
    "H=scan_mode" & vbCrLf & _
    "H=<br>" & vbCrLf & _
    "H=software_version" & vbCrLf & _
    "H=software_pluginscount" & vbCrLf & _
    "H=<br>"

    report_vulnerability_structure = report_vulnerability_structure & _
    "V=plugin_id" & vbCrLf & _
    "V=plugin_name" & vbCrLf & _
    "V=plugin_protocol" & vbCrLf & _
    "V=plugin_port" & vbCrLf & _
    "V=bug_severity" & vbCrLf & _
    "V=bug_advisory" & vbCrLf & _
    "V=bug_affected" & vbCrLf & _
    "V=bug_not_affected" & vbCrLf & _
    "V=bug_vulnerability_class" & vbCrLf & _
    "V=bug_exploit_url" & vbCrLf
    
    report_vulnerability_structure = report_vulnerability_structure & _
    "V=<br>" & vbCrLf & _
    "V=bug_description" & vbCrLf & _
    "V=<br>" & vbCrLf & _
    "V=bug_response" & vbCrLf & _
    "V=<br>" & vbCrLf & _
    "V=bug_solution" & vbCrLf & _
    "V=<br>" & vbCrLf & _
    "V=source_cve" & vbCrLf & _
    "V=source_securityfocus_bid"
End Function

Public Function PrepareReportStructure()
    Dim i As Integer
    Dim HeaderListCount As Integer
    Dim VulnerabilityListCount As Integer

    'Delete the old report template content
    report_header_structure = vbNullString
    report_vulnerability_structure = vbNullString

    HeaderListCount = frmReportConfiguration.lstHeaderReport.ListCount - 1
    VulnerabilityListCount = frmReportConfiguration.lstVulnerabilityReport.ListCount - 1

    For i = 0 To HeaderListCount
        report_header_structure = report_header_structure & _
            "H=" & frmReportConfiguration.lstHeaderReport.List(i) & vbCrLf
    Next i

    For i = 0 To VulnerabilityListCount
        report_vulnerability_structure = report_vulnerability_structure & _
            "V=" & frmReportConfiguration.lstVulnerabilityReport.List(i) & vbCrLf
    Next i
End Function

Public Sub WriteReportTemplateToFile()
    'save the actual report template as file
    On Error Resume Next 'Workaround! No directory checking yet!
    Open App.Path & "\reporttemplates\default.reporttemplate" For Output As #1
        Print #1, report_header_structure & report_vulnerability_structure
    Close
End Sub

Private Sub WriteDefaultReportToFile()
    'save default report to a file
    On Error Resume Next 'Workaround! No directory checking yet!
    Open App.Path & "\reporttemplates\default.reporttemplate" For Output As #1
        Print #1, report_header_structure & vbCrLf & report_vulnerability_structure
    Close
End Sub

Public Sub WriteReportHeader()
    Dim ReportHeaderStructureArray() As String
    Dim ReportHeaderStructureArrayCount As Integer
    Dim i As Integer

    ReportHeaderStructureArray = Split(report_header_structure, vbCrLf)
    ReportHeaderStructureArrayCount = UBound(ReportHeaderStructureArray)

    'Open the default report
    For i = 0 To ReportHeaderStructureArrayCount
        'write the selected item
        If InStr(3, ReportHeaderStructureArray(i), "scan_target") Then
            ActualReport = ActualReport & _
                "Scan target: " & Target & vbCrLf
        ElseIf InStr(3, ReportHeaderStructureArray(i), "scan_date") Then
            ActualReport = ActualReport & _
                "Scan date: " & Date & vbCrLf
        ElseIf InStr(3, ReportHeaderStructureArray(i), "scan_time") Then
            ActualReport = ActualReport & _
                "Scan time: " & Time & vbCrLf
        ElseIf InStr(3, ReportHeaderStructureArray(i), "scan_mode") Then
            ActualReport = ActualReport & _
                "Attack mode: " & AttackMode & vbCrLf
        ElseIf InStr(3, ReportHeaderStructureArray(i), "software_version") Then
            ActualReport = ActualReport & _
                "Software: " & frmMain.Caption & vbCrLf
        ElseIf InStr(3, ReportHeaderStructureArray(i), "software_pluginscount") Then
            ActualReport = ActualReport & _
                "Plugins count: " & HowManyLoadedPlugins & vbCrLf
        ElseIf InStr(3, ReportHeaderStructureArray(i), "<br>") Then
           ActualReport = ActualReport & vbCrLf
        End If
    Next i
End Sub

Public Sub WriteReportVulnerability()
    Dim ReportVulnerabilityStructureArray() As String
    Dim ReportVulnerabilityStructureArrayCount As Integer
    Dim i As Integer

    ReportVulnerabilityStructureArray = Split(report_vulnerability_structure, vbCrLf)
    ReportVulnerabilityStructureArrayCount = UBound(ReportVulnerabilityStructureArray)
    

    'Open and read the report template file
    For i = 0 To ReportVulnerabilityStructureArrayCount
        'write the selected item
        If InStr(3, ReportVulnerabilityStructureArray(i), "plugin_id") Then
            ActualReport = ActualReport & _
                "Plugin ID: " & plugin_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_name") Then
            ActualReport = ActualReport & _
                "Plugin name: " & plugin_name & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_filename") Then
            ActualReport = ActualReport & _
                "Plugin filename: " & plugin_filename & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_filesize") Then
            ActualReport = ActualReport & _
                "Plugin filesize: " & plugin_filesize & " bytes" & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_family") Then
            ActualReport = ActualReport & _
                "Plugin family: " & plugin_family & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_created_name") Then
            ActualReport = ActualReport & _
                "Plugin created name: " & plugin_created_name & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_created_email") Then
            ActualReport = ActualReport & _
                "Plugin created email: " & plugin_created_email & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_created_web") Then
            ActualReport = ActualReport & _
                "Plugin created web: " & plugin_created_web & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_created_company") Then
            ActualReport = ActualReport & _
                "Plugin created company: " & plugin_created_company & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_created_date") Then
            ActualReport = ActualReport & _
                "Plugin created date: " & plugin_created_date & vbCrLf
        
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_updated_name") Then
            ActualReport = ActualReport & _
                "Plugin updated name: " & plugin_updated_name & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_updated_email") Then
            ActualReport = ActualReport & _
                "Plugin updated email: " & plugin_updated_email & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_updated_web") Then
            ActualReport = ActualReport & _
                "Plugin updated web: " & plugin_updated_web & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_updated_company") Then
            ActualReport = ActualReport & _
                "Plugin updated company: " & plugin_updated_company & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_updated_date") Then
            ActualReport = ActualReport & _
                "Plugin updated date: " & plugin_updated_date & vbCrLf
        
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_version") Then
            ActualReport = ActualReport & _
                "Plugin version: " & plugin_version & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_changelog") Then
            ActualReport = ActualReport & _
                "Plugin changelog: " & plugin_changelog & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_protocol") Then
            ActualReport = ActualReport & _
                "Protocol: " & plugin_protocol & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_port") Then
            ActualReport = ActualReport & _
                "Port: " & plugin_port & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_procedure_detection") Then
            ActualReport = ActualReport & _
                "Plugin procedure detection: " & plugin_procedure_detection & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_procedure_exploit") Then
            ActualReport = ActualReport & _
                "Plugin procedure exploit: " & plugin_procedure_exploit & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "plugin_comment") Then
            ActualReport = ActualReport & _
                "Plugin comment: " & plugin_comment & vbCrLf
''''''''''''''''''''''''''''''''''''''''''
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_published_name") Then
            ActualReport = ActualReport & _
                "Bug published name: " & bug_published_name & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_published_email") Then
            ActualReport = ActualReport & _
                "Bug published email: " & bug_published_email & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_published_web") Then
            ActualReport = ActualReport & _
                "Bug published web: " & bug_published_web & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_published_company") Then
            ActualReport = ActualReport & _
                "Bug published company: " & bug_published_company & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_published_date") Then
            ActualReport = ActualReport & _
                "Bug published date: " & bug_published_date & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_advisory") Then
            ActualReport = ActualReport & _
                "Advisory: " & bug_advisory & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_produced_name") Then
            ActualReport = ActualReport & _
                "Bug produced name: " & bug_produced_name & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_produced_email") Then
            ActualReport = ActualReport & _
                "Bug produced email: " & bug_produced_email & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_produced_web") Then
            ActualReport = ActualReport & _
                "Bug produced web: " & bug_produced_web & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_affected") Then
            ActualReport = ActualReport & _
                "Affected: " & bug_affected & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_not_affected") Then
            ActualReport = ActualReport & _
                "Not affected: " & bug_not_affected & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_false_positives") Then
            ActualReport = ActualReport & _
                "False positives: " & bug_false_positives & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_false_negatives") Then
            ActualReport = ActualReport & _
                "False negatives: " & bug_false_negatives & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_vulnerability_class") Then
            ActualReport = ActualReport & _
                "Vulnerability class: " & bug_vulnerability_class & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_description") Then
            ActualReport = ActualReport & _
                "Description: " & bug_description & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_response") Then
            ActualReport = ActualReport & _
                "Response: " & vbCrLf & vbCrLf & _
                    Mid(LastResponse, 1, 1000)
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_solution") Then
            ActualReport = ActualReport & _
                "Solution: " & bug_solution & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_fixing_time") Then
            ActualReport = ActualReport & _
                "Time to fix: " & bug_fixing_time & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_exploit_availability") Then
            ActualReport = ActualReport & _
                "Exploit available: " & bug_exploit_availability & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_exploit_url") Then
            ActualReport = ActualReport & _
                "Exploit URL: " & bug_exploit_url & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_remote") Then
            ActualReport = ActualReport & _
                "Bug remote: " & bug_remote & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_local") Then
            ActualReport = ActualReport & _
                "Bug local: " & bug_local & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_severity") Then
            ActualReport = ActualReport & _
                "Severity: " & bug_severity & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_popularity") Then
            ActualReport = ActualReport & _
                "Popularity: " & bug_popularity & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_simplicity") Then
            ActualReport = ActualReport & _
                "Simplicity: " & bug_simplicity & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_impact") Then
            ActualReport = ActualReport & _
                "Impact: " & bug_impact & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_risk") Then
            ActualReport = ActualReport & _
                "Risk: " & bug_risk & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_nessus_risk") Then
            ActualReport = ActualReport & _
                "Nessus risk: " & bug_nessus_risk & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_iss_scanner_rating") Then
            ActualReport = ActualReport & _
                "ISS Scanner rating: " & bug_iss_scanner_rating & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_netrecon_rating") Then
            ActualReport = ActualReport & _
                "Symantec NetRecon rating: " & bug_netrecon_rating & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "bug_check_tool") Then
            ActualReport = ActualReport & _
                "Tools for further checks: " & bug_check_tool & vbCrLf
''''''''''''''''''''''''''''''''''''''''''
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_cve") Then
            ActualReport = ActualReport & _
                "CVE: " & source_cve & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_certvu_id") Then
            ActualReport = ActualReport & _
                "CERT Vulnerability Note ID: " & source_certvu_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_cert_id") Then
            ActualReport = ActualReport & _
                "CERT ID: " & source_cert_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_uscertta_id") Then
            ActualReport = ActualReport & _
                "US-CERT ID: " & source_uscertta_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_securityfocus_bid") Then
            ActualReport = ActualReport & _
                "SecurityFocus BID: " & source_securityfocus_bid & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_osvdb_id") Then
            ActualReport = ActualReport & _
                "OSVDB ID: " & source_osvdb_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_secunia_id") Then
            ActualReport = ActualReport & _
                "Secunia ID: " & source_secunia_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_securiteam_url") Then
            ActualReport = ActualReport & _
                "SecuriTeam URL: " & source_securiteam_url & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_securitytracker_id") Then
            ActualReport = ActualReport & _
                "Security Tracker ID: " & source_securitytracker_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_scip_id") Then
            ActualReport = ActualReport & _
                "scipID: " & source_scip_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_tecchannel_id") Then
            ActualReport = ActualReport & _
                "tecchannel ID: " & source_tecchannel_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_heise_news") Then
            ActualReport = ActualReport & _
                "Heise News: " & source_heise_news & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_heise_security") Then
            ActualReport = ActualReport & _
                "Heise Security: " & source_heise_security & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_aerasec_id") Then
            ActualReport = ActualReport & _
                "AeraSec ID: " & source_aerasec_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_nessus_id") Then
            ActualReport = ActualReport & _
                "Nessus ID: " & source_nessus_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_issxforce_id") Then
            ActualReport = ActualReport & _
                "ISS X-Force ID: " & source_issxforce_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_snort_id") Then
           ActualReport = ActualReport & _
                "Snort ID: " & source_snort_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_arachnids_id") Then
           ActualReport = ActualReport & _
                "ArachnIDS ID: " & source_arachnids_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_mssb_id") Then
           ActualReport = ActualReport & _
                "Microsoft Security Bulletin ID: " & source_mssb_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_mskb_id") Then
           ActualReport = ActualReport & _
                "Microsoft Knowledge Base ID: " & source_mskb_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_netbsdsa_id") Then
           ActualReport = ActualReport & _
                "NetBSD Security Advisory ID: " & source_netbsdsa_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_rhsa_id") Then
           ActualReport = ActualReport & _
                "RedHat Security Advisory ID: " & source_rhsa_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_ciac_id") Then
           ActualReport = ActualReport & _
                "CIAC ID: " & source_ciac_id & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_literature") Then
           ActualReport = ActualReport & _
                "Literature: " & source_literature & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "source_misc") Then
           ActualReport = ActualReport & _
                "Misc sources: " & source_misc & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "report_header_structure") Then
           ActualReport = ActualReport & _
                "Report header structure: " & report_header_structure & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "report_vulnerability_structure") Then
           ActualReport = ActualReport & _
                "Report vulnerability structure: " & report_vulnerability_structure & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "<br>") Then
           ActualReport = ActualReport & vbCrLf
''''''''''''''''''''''''''''''''
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "session_procedure_type") Then
            ActualReport = ActualReport & _
                "Session procedure type: " & session_procedure_type & vbCrLf
        ElseIf InStr(3, ReportVulnerabilityStructureArray(i), "session_procedure_commands") Then
            ActualReport = ActualReport & _
                "Session procedure commands: " & session_procedure_commands & vbCrLf
        End If
    Next i
    
    'Show the last response data if there is some
    If IsFormVisible("frmReport") = True Then
        frmReport.txtReport.Text = ActualReport
    End If
End Sub

Public Function PluginReportData() As String
    Dim TempReport As String    'Here we chache the actual report
    Dim report_header_structure_backup As String
    Dim report_vulnerability_structure_backup As String
    
    report_header_structure_backup = report_header_structure
    report_vulnerability_structure_backup = report_vulnerability_structure
    
    'Write the new report in the file
    Call PrepareReportStructure
    
    'Write the actual report into the temp report
    TempReport = ActualReport
    
    'Clear the actual report
    ActualReport = vbNullString
    
    'Compute the report
    Call WriteReportHeader
    Call WriteReportVulnerability
    
    PluginReportData = ActualReport
    
    'Write the old report back into the actual report
    ActualReport = TempReport
End Function

'Public Sub WritePluginNameToReportFile(ByRef InputString As String)
'    'Write the collected data into the file; the plugin name will be the file name
'    Open ReportsDirectory & "\" & Target & ".report" For Append As 1
'        Print #1, InputString
'    Close
'End Sub
