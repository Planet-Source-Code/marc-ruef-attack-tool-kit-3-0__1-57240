Attribute VB_Name = "modPluginHandling"
' This module reads the plugin, performs the parsing and writes the result
' in the global variables.

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 3.0 2004-11-13                                                           *
' * - Introduced the plugin_changelog, bug_false_positives and _negatives fields.    *
' * Version 3.0 2004-11-01                                                           *
' * - Replaced all useless functions with normal subs.                               *
' * Version 3.0 2004-10-01                                                           *
' * - Added the session variants.                                                    *
' * Version 2.1 2004-09-08                                                           *
' * - Additional filling of the fields source_misc and source_literature if empty.   *
' * Version 2.0 2004-08-24                                                           *
' * - Changed Len to LenB checking during ATK plugin parsing. This increases the     *
' *   speed of the procedure very much.                                              *
' * - Increased the speed of some array handling during the writing of a plugin.     *
' * Version 2.0 2004-08-16                                                           *
' * - Corrected an error with the closing tag bug_not_affected during writing.       *
' ************************************************************************************

Option Explicit

                                            'In this column is a short description
                                            'of the meaning of the variables. You
                                            'find more information about them on the
                                            'project web site or the readme.

Public plugin_filename As String            'The filename of the plugin. This value is not
                                            'saved in the plugin file.
Public plugin_filesize As String            'The filesize of the plugin. Also not a saved
                                            'value in the plugin file.

Public plugin_id As String                  'Unique ID of the ATK plugin
Public plugin_name As String                'Plugin name of the ATK plugin
Public plugin_family As String              'Plugin family of the ATK plugin
Public plugin_created_name As String        'Name of the person who created the plugin
Public plugin_created_email As String       'Email of the person who created the plugin
Public plugin_created_web As String         'Web site of the person who created the plugin
Public plugin_created_company As String     'Companyname of the person who created the plugin
Public plugin_created_date As String        'When was the ATK plugin created
Public plugin_updated_name As String        'Name of the person who updated the ATK plugin
Public plugin_updated_email As String       'Email of the person who updated the ATK plugin
Public plugin_updated_web As String         'Web site of the person who updated the ATK plugin
Public plugin_updated_company As String     'Company name of the person who updated the ATK plugin
Public plugin_updated_date As String        'When was the ATK plugin updated the last time
Public plugin_version As String             'Which version of the plugin is it
Public plugin_changelog As String           'The changelog of the plugin
Public plugin_protocol As String            'Which protocol does the plugin use
Public plugin_port As String                'Which port does the ATK plugin use
Public plugin_procedure_detection As String 'The request procedure for detection
Public plugin_procedure_exploit As String   'The request procedure for exploit
Public plugin_detection_accuracy As String  'The accuracy for detection
Public plugin_exploit_accuracy As String    'The accuracy for exploiting
Public plugin_comment As String             'Some words about the ATK plugin

Public bug_published_name As String         'Who published the bug first
Public bug_published_email As String        'Who published the bug first
Public bug_published_web As String          'Who published the bug first
Public bug_published_company As String      'Who published the bug first
Public bug_published_date As String         'Who published the bug first
Public bug_produced_name As String          'Who produced the product
Public bug_produced_email As String         'Who produced the product
Public bug_produced_web As String           'Who produced the product
Public bug_advisory As String               'What is the name and URL of the advisory
Public bug_affected As String               'Which systems and solutions are affected
Public bug_not_affected As String           'Which systems and solutions are not affected
Public bug_false_positives As String        'Known false-positives
Public bug_false_negatives As String        'Known false-negatives
Public bug_vulnerability_class As String    'The class of the vulnerability
Public bug_local As String                  'A local vulnerability
Public bug_remote As String                 'A remote vulnerability
Public bug_description As String            'The description of the vulnerability
Public bug_solution As String               'The solution(s) for the vulnerability
Public bug_fixing_time As String            'The time needed to fix the bug (e.g. hours)
Public bug_exploit_availability As String   'The existence of an exploit
Public bug_exploit_url As String            'The URL of the exploit
Public bug_severity As String               'The severity of the vulnerability
Public bug_popularity As String             'The popularity of the vulnerability (1 to 10)
Public bug_simplicity As String             'The simplicity of the bug (1 to 10)
Public bug_impact As String                 'The impact level of the bug (1 to 10)
Public bug_risk As String                   'The rist of the bug (1 to 10)
Public bug_nessus_risk As String            'The risk level of the bug by Nessus
Public bug_iss_scanner_rating As String     'The risk level by ISS Scanners
Public bug_netrecon_rating As String        'The risk level by Symantec NetRecon
Public bug_check_tool As String             'List of tools that are able to check the bug

Public source_cve As String                 'The unique CAN or CVE number of the vulnerability
Public source_certvu_id As String           'The unique CERT Vulnerability ID
Public source_cert_id As String             'The unique CERT ID
Public source_uscertta_id As String         'The unique US-CERT Technical Advisory ID
Public source_securityfocus_bid As String   'The unique SecurityFocus/Bugtraq ID
Public source_osvdb_id As String            'The unique Open Source Vulnerability Data Base ID
Public source_secunia_id As String          'The unique Secunia ID of the vulnerability
Public source_securiteam_url As String      'The SecuriTeam.com URL
Public source_securitytracker_id As String  'The SecurityTracker ID
Public source_scip_id As String             'The unique scipID of the vulnerability
Public source_tecchannel_id As String       'The unique tecchannel ID
Public source_heise_news As String          'The unique Heise News
Public source_heise_security As String      'The unique Heise Security
Public source_aerasec_id As String          'The unique AeraSec ID
Public source_nessus_id As String           'The unique Nessus ID of the vulnerability
Public source_issxforce_id As String        'The unique ISS X-Force ID
Public source_snort_id As String            'The unique Snort ID of the vulnerability
Public source_arachnids_id As String        'The unique ArachnIDS ID
Public source_mssb_id As String             'The unique Microsoft Security Bulletin ID
Public source_mskb_id As String             'The unique Microsoft Knowledge-Base Article ID
Public source_netbsdsa_id As String         'The unique NetBSD Security Advisory ID
Public source_rhsa_id As String             'The unique Red Hat Security Advisory ID
Public source_ciac_id As String             'The unique CIAC ID
Public source_literature As String          'List of books about the flaw
Public source_misc As String                'List of other sources (e.g. TV shows)

Public session_procedure_type As String     'Type of the procedure (detection or exploit)
Public session_procedure_commands As String 'The commands of the defined procedure
Public session_triggers As String           'The triggers of the check and type
Public session_trigger_match As String      'The matching trigger

' *******************************************************************
' * Reset all plugin variables. This is usually done before new     *
' * data is read. Old "garbage" is prevented and the software don't *
' * need so much ressources during runtime.                         *
' *******************************************************************

Public Sub ClearAllPluginVariables()
    If frmMain.lblVulnerabilityState.BackColor <> &HE0E0E0 Then
        Dim strAlertingText As String
        
        strAlertingText = "The vulnerability was not tested. " & _
            "Please run the selected plugin to determine the existence of the flaw."
        
        'Message if the vulnerability was found
        frmMain.lblVulnerabilityState.Caption = strAlertingText
        frmMain.lblVulnerabilityState.BackColor = &HE0E0E0
    End If
    
    plugin_id = vbNullString
    plugin_name = vbNullString
    plugin_family = vbNullString
    plugin_created_name = vbNullString
    plugin_created_email = vbNullString
    plugin_created_web = vbNullString
    plugin_created_company = vbNullString
    plugin_created_date = vbNullString
    plugin_updated_name = vbNullString
    plugin_updated_email = vbNullString
    plugin_updated_web = vbNullString
    plugin_updated_company = vbNullString
    plugin_updated_date = vbNullString
    plugin_version = vbNullString
    plugin_changelog = vbNullString
    plugin_protocol = vbNullString
    plugin_port = vbNullString
    plugin_procedure_detection = vbNullString
    plugin_procedure_exploit = vbNullString
    plugin_detection_accuracy = vbNullString
    plugin_exploit_accuracy = vbNullString
    plugin_comment = vbNullString
    
    bug_published_name = vbNullString
    bug_published_email = vbNullString
    bug_published_web = vbNullString
    bug_published_company = vbNullString
    bug_published_date = vbNullString
    bug_produced_name = vbNullString
    bug_produced_email = vbNullString
    bug_produced_web = vbNullString
    bug_advisory = vbNullString
    bug_affected = vbNullString
    bug_not_affected = vbNullString
    bug_false_positives = vbNullString
    bug_false_negatives = vbNullString
    bug_vulnerability_class = vbNullString
    bug_local = vbNullString
    bug_remote = vbNullString
    bug_description = vbNullString
    bug_solution = vbNullString
    bug_fixing_time = vbNullString
    bug_exploit_availability = vbNullString
    bug_exploit_url = vbNullString
    bug_severity = vbNullString
    bug_popularity = vbNullString
    bug_simplicity = vbNullString
    bug_impact = vbNullString
    bug_risk = vbNullString
    bug_nessus_risk = vbNullString
    bug_iss_scanner_rating = vbNullString
    bug_netrecon_rating = vbNullString
    bug_check_tool = vbNullString
    
    source_cve = vbNullString
    source_certvu_id = vbNullString
    source_cert_id = vbNullString
    source_uscertta_id = vbNullString
    source_securityfocus_bid = vbNullString
    source_osvdb_id = vbNullString
    source_secunia_id = vbNullString
    source_securiteam_url = vbNullString
    source_securitytracker_id = vbNullString
    source_scip_id = vbNullString
    source_tecchannel_id = vbNullString
    source_heise_news = vbNullString
    source_heise_security = vbNullString
    source_aerasec_id = vbNullString
    source_nessus_id = vbNullString
    source_issxforce_id = vbNullString
    source_snort_id = vbNullString
    source_arachnids_id = vbNullString
    source_mssb_id = vbNullString
    source_mskb_id = vbNullString
    source_netbsdsa_id = vbNullString
    source_rhsa_id = vbNullString
    source_ciac_id = vbNullString
    source_literature = vbNullString
    source_misc = vbNullString
    
    session_procedure_type = vbNullString
    session_procedure_commands = vbNullString
    session_triggers = vbNullString
    session_trigger_match = vbNullString
End Sub

Public Function ReadPluginFromFile(ByRef Filename As String) As String
    Dim Temp As String          'The temporary file output

    'Check the existence of the file
    On Error Resume Next
    If Len(Dir(PluginDirectory & "\" & Filename)) > 1 Then
        plugin_filename = Filename
        
        'Open and read the plugin file
        Open PluginDirectory & "\" & Filename For Input As 1
            Do While Not EOF(1)
                Line Input #1, Temp
                    ReadPluginFromFile = ReadPluginFromFile & Temp
            Loop
        Close
    
        'Set the plugin silesize
        plugin_filesize = Len(ReadPluginFromFile)
    Else
        Call errPluginDoesNotExist(Filename)
    End If
End Function

Public Sub ParseATKPlugin(ByRef ATKPluginContent As String)
    Dim TempArray() As String       'A temporary array for the splitting and parsing
    Dim strErrorousField As String  'Here we save the name of the field that has an error
    
    'Reset the error field
    strErrorousField = vbNullString
    
    'Prevent error messages if a field does not exist
    On Error Resume Next
    
    'Clear the values from the last plugin to prevent misunderstandings
    Call ClearAllPluginVariables    'Plugin variables itself
    'Call ClearAllResponseVariables  'Plugin last response
    
    'Get the data fields and write them into the public variables
    TempArray = Split(ATKPluginContent, "<plugin_id>")
    TempArray = Split(TempArray(1), "</plugin_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_id = TempArray(0)
    Else
        strErrorousField = "plugin_id"
    End If
        
    TempArray = Split(ATKPluginContent, "<plugin_name>")
    TempArray = Split(TempArray(1), "</plugin_name>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_name = TempArray(0)
    Else
        strErrorousField = "plugin_name"
    End If

    TempArray = Split(ATKPluginContent, "<plugin_family>")
    TempArray = Split(TempArray(1), "</plugin_family>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_family = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_created_name>")
    TempArray = Split(TempArray(1), "</plugin_created_name>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_created_name = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_created_email>")
    TempArray = Split(TempArray(1), "</plugin_created_email>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_created_email = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_created_web>")
    TempArray = Split(TempArray(1), "</plugin_created_web>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_created_web = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_created_company>")
    TempArray = Split(TempArray(1), "</plugin_created_company>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_created_company = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_created_date>")
    TempArray = Split(TempArray(1), "</plugin_created_date>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_created_date = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_updated_name>")
    TempArray = Split(TempArray(1), "</plugin_updated_name>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_updated_name = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_updated_email>")
    TempArray = Split(TempArray(1), "</plugin_updated_email>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_updated_email = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_updated_web>")
    TempArray = Split(TempArray(1), "</plugin_updated_web>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_updated_web = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_updated_company>")
    TempArray = Split(TempArray(1), "</plugin_updated_company>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_updated_company = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<plugin_updated_date>")
    TempArray = Split(TempArray(1), "</plugin_updated_date>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_updated_date = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_version>")
    TempArray = Split(TempArray(1), "</plugin_version>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_version = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_changelog>")
    TempArray = Split(TempArray(1), "</plugin_changelog>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_changelog = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_protocol>")
    TempArray = Split(TempArray(1), "</plugin_protocol>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_protocol = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<plugin_port>")
    TempArray = Split(TempArray(1), "</plugin_port>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_port = Val(TempArray(0))
    Else
        strErrorousField = "plugin_port"
    End If
    
    TempArray = Split(ATKPluginContent, "<plugin_procedure_detection>")
    TempArray = Split(TempArray(1), "</plugin_procedure_detection>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_procedure_detection = TempArray(0)
        frmMain.mnuPluginsRunDetectionItem.Enabled = True
    Else
        frmMain.mnuPluginsRunDetectionItem.Enabled = False
    End If
                
    TempArray = Split(ATKPluginContent, "<plugin_procedure_exploit>")
    TempArray = Split(TempArray(1), "</plugin_procedure_exploit>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_procedure_exploit = TempArray(0)
        frmMain.mnuPluginsRunExploitItem.Enabled = True
    Else
        frmMain.mnuPluginsRunExploitItem.Enabled = False
    End If
                
    TempArray = Split(ATKPluginContent, "<plugin_detection_accuracy>")
    TempArray = Split(TempArray(1), "</plugin_detection_accuracy>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_detection_accuracy = TempArray(0)
    End If
                
    TempArray = Split(ATKPluginContent, "<plugin_exploit_accuracy>")
    TempArray = Split(TempArray(1), "</plugin_exploit_accuracy>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_exploit_accuracy = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<plugin_comment>")
    TempArray = Split(TempArray(1), "</plugin_comment>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        plugin_comment = TempArray(0)
    End If
                
    TempArray = Split(ATKPluginContent, "<bug_published_name>")
    TempArray = Split(TempArray(1), "</bug_published_name>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_published_name = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_published_email>")
    TempArray = Split(TempArray(1), "</bug_published_email>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_published_email = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_published_web>")
    TempArray = Split(TempArray(1), "</bug_published_web>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_published_web = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_published_company>")
    TempArray = Split(TempArray(1), "</bug_published_company>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_published_company = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_published_date>")
    TempArray = Split(TempArray(1), "</bug_published_date>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_published_date = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_advisory>")
    TempArray = Split(TempArray(1), "</bug_advisory>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_advisory = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_produced_name>")
    TempArray = Split(TempArray(1), "</bug_produced_name>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_produced_name = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_produced_email>")
    TempArray = Split(TempArray(1), "</bug_produced_email>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_produced_email = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_produced_web>")
    TempArray = Split(TempArray(1), "</bug_produced_web>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_produced_web = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_affected>")
    TempArray = Split(TempArray(1), "</bug_affected>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_affected = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_not_affected>")
    TempArray = Split(TempArray(1), "</bug_not_affected>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_not_affected = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_false_positives>")
    TempArray = Split(TempArray(1), "</bug_false_positives>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_false_positives = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_false_negatives>")
    TempArray = Split(TempArray(1), "</bug_false_negatives>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_false_negatives = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_local>")
    TempArray = Split(TempArray(1), "</bug_local>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_local = TempArray(0)
    End If
        
    TempArray = Split(ATKPluginContent, "<bug_remote>")
    TempArray = Split(TempArray(1), "</bug_remote>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_remote = TempArray(0)
    End If
        
    TempArray = Split(ATKPluginContent, "<bug_vulnerability_class>")
    TempArray = Split(TempArray(1), "</bug_vulnerability_class>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_vulnerability_class = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_description>")
    TempArray = Split(TempArray(1), "</bug_description>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_description = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_solution>")
    TempArray = Split(TempArray(1), "</bug_solution>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_solution = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_fixing_time>")
    TempArray = Split(TempArray(1), "</bug_fixing_time>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_fixing_time = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_exploit_availability>")
    TempArray = Split(TempArray(1), "</bug_exploit_availability>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_exploit_availability = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_exploit_url>")
    TempArray = Split(TempArray(1), "</bug_exploit_url>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_exploit_url = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<bug_severity>")
    TempArray = Split(TempArray(1), "</bug_severity>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_severity = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<bug_popularity>")
    TempArray = Split(TempArray(1), "</bug_popularity>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_popularity = Val(TempArray(0))
    End If

    TempArray = Split(ATKPluginContent, "<bug_simplicity>")
    TempArray = Split(TempArray(1), "</bug_simplicity>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_simplicity = Val(TempArray(0))
    End If

    TempArray = Split(ATKPluginContent, "<bug_impact>")
    TempArray = Split(TempArray(1), "</bug_impact>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_impact = Val(TempArray(0))
    End If

    TempArray = Split(ATKPluginContent, "<bug_risk>")
    TempArray = Split(TempArray(1), "</bug_risk>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_risk = Val(TempArray(0))
    End If

    TempArray = Split(ATKPluginContent, "<bug_nessus_risk>")
    TempArray = Split(TempArray(1), "</bug_nessus_risk>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_nessus_risk = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<bug_iss_scanner_rating>")
    TempArray = Split(TempArray(1), "</bug_iss_scanner_rating>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_iss_scanner_rating = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<bug_netrecon_rating>")
    TempArray = Split(TempArray(1), "</bug_netrecon_rating>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_netrecon_rating = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<bug_check_tool>")
    TempArray = Split(TempArray(1), "</bug_check_tool>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        bug_check_tool = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<source_cve>")
    TempArray = Split(TempArray(1), "</source_cve>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_cve = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<source_certvu_id>")
    TempArray = Split(TempArray(1), "</source_certvu_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_certvu_id = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<source_cert_id>")
    TempArray = Split(TempArray(1), "</source_cert_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_cert_id = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<source_uscertta_id>")
    TempArray = Split(TempArray(1), "</source_uscertta_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_uscertta_id = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<source_securityfocus_bid>")
    TempArray = Split(TempArray(1), "</source_securityfocus_bid>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_securityfocus_bid = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_osvdb_id>")
    TempArray = Split(TempArray(1), "</source_osvdb_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_osvdb_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_secunia_id>")
    TempArray = Split(TempArray(1), "</source_secunia_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_secunia_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_securiteam_url>")
    TempArray = Split(TempArray(1), "</source_securiteam_url>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_securiteam_url = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_securitytracker_id>")
    TempArray = Split(TempArray(1), "</source_securitytracker_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_securitytracker_id = TempArray(0)
    End If
                
    TempArray = Split(ATKPluginContent, "<source_scip_id>")
    TempArray = Split(TempArray(1), "</source_scip_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_scip_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_tecchannel_id>")
    TempArray = Split(TempArray(1), "</source_tecchannel_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_tecchannel_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_heise_news>")
    TempArray = Split(TempArray(1), "</source_heise_news>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_heise_news = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_heise_security>")
    TempArray = Split(TempArray(1), "</source_heise_security>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_heise_security = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_aerasec_id>")
    TempArray = Split(TempArray(1), "</source_aerasec_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_aerasec_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_nessus_id>")
    TempArray = Split(TempArray(1), "</source_nessus_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_nessus_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_issxforce_id>")
    TempArray = Split(TempArray(1), "</source_issxforce_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_issxforce_id = TempArray(0)
    End If
        
    TempArray = Split(ATKPluginContent, "<source_snort_id>")
    TempArray = Split(TempArray(1), "</source_snort_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_snort_id = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<source_arachnids_id>")
    TempArray = Split(TempArray(1), "</source_arachnids_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_arachnids_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_mssb_id>")
    TempArray = Split(TempArray(1), "</source_mssb_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_mssb_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_mskb_id>")
    TempArray = Split(TempArray(1), "</source_mskb_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_mskb_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_netbsdsa_id>")
    TempArray = Split(TempArray(1), "</source_netbsdsa_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_netbsdsa_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_rhsa_id>")
    TempArray = Split(TempArray(1), "</source_rhsa_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_rhsa_id = TempArray(0)
    End If
    
    TempArray = Split(ATKPluginContent, "<source_ciac_id>")
    TempArray = Split(TempArray(1), "</source_ciac_id>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_ciac_id = TempArray(0)
    End If
     
    TempArray = Split(ATKPluginContent, "<source_literature>")
    TempArray = Split(TempArray(1), "</source_literature>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_literature = TempArray(0)
    End If

    TempArray = Split(ATKPluginContent, "<source_misc>")
    TempArray = Split(TempArray(1), "</source_misc>")
    If LenB(TempArray(0)) <> LenB(ATKPluginContent) Then
        source_misc = TempArray(0)
    End If

    If LenB(strErrorousField) <> 0 Then
        Call errPluginDataMissing(strErrorousField, plugin_filename, plugin_id)
    End If
End Sub

Private Sub PluginReadError(ByRef Filename As String, ByRef Position As String)
    'Write error message if something went wrong during parsing of the plugin file
    MsgBox ("Could not find the data field '" & Position & "'" & vbCrLf & _
        "in the plugin '" & Filename & "'." & vbCrLf & vbCrLf & _
        "The plugin seems to be broken and can't be used." & vbCrLf & _
        "Please check this manually."), _
        vbInformation, "Attack Tool Kit Plugin parsing error"
End Sub

Public Sub WritePluginToFile(ByRef Filename As String)
    Dim PluginContent As String
    Dim PluginContentArray() As String
    Dim PluginContentItemCount As Integer
    Dim i As Integer
    
    'Prepare the comment
    If LenB(plugin_comment) = 0 Then
        plugin_comment = "This plugin was written with the ATK Attack Editor."
    End If
    
    'Prepare the exploit URL if a SecurityFocus exploit may be given
    If LenB(bug_exploit_url) = 0 Then
        If LenB(source_securityfocus_bid) <> 0 Then
            bug_exploit_url = "http://www.securityfocus.com/bid/" & source_securityfocus_bid & "/exploit/"
        End If
    End If
    
    'Prepare the misc source
    If LenB(source_misc) = 0 Then
        source_misc = "http://www.computec.ch"
    End If
    
    'Prepare the literature
    If LenB(source_literature) = 0 Then
        source_literature = "Hacking Intern - Angriffe, Strategien, Abwehr, " & _
        "Marc Ruef, Marko Rogge, Uwe Velten and Wolfram Gieseke, " & _
        "November 1, 2002, Data Becker, Düsseldorf, ISBN 381582284X"
    End If
    
    'Collect the whole data
    'Add the plugin data
    PluginContent = _
        "<plugin_id>" & plugin_id & "</plugin_id>" & vbCrLf & _
        "<plugin_name>" & plugin_name & "</plugin_name>" & vbCrLf & _
        "<plugin_family>" & plugin_family & "</plugin_family>" & vbCrLf & _
        "<plugin_created_date>" & plugin_created_date & "</plugin_created_date>" & vbCrLf & _
        "<plugin_created_name>" & plugin_created_name & "</plugin_created_name>" & vbCrLf & _
        "<plugin_created_email>" & plugin_created_email & "</plugin_created_email>" & vbCrLf & _
        "<plugin_created_web>" & plugin_created_web & "</plugin_created_web>" & vbCrLf & _
        "<plugin_created_company>" & plugin_created_company & "</plugin_created_company>" & vbCrLf & _
        "<plugin_updated_name>" & plugin_updated_name & "</plugin_updated_name>" & vbCrLf & _
        "<plugin_updated_email>" & plugin_updated_email & "</plugin_updated_email>" & vbCrLf & _
        "<plugin_updated_web>" & plugin_updated_web & "</plugin_updated_web>" & vbCrLf & _
        "<plugin_updated_company>" & plugin_updated_company & "</plugin_updated_company>" & vbCrLf & _
        "<plugin_updated_date>" & plugin_updated_date & "</plugin_updated_date>" & vbCrLf & _
        "<plugin_version>" & plugin_version & "</plugin_version>" & vbCrLf & _
        "<plugin_changelog>" & plugin_changelog & "</plugin_changelog>" & vbCrLf & _
        "<plugin_protocol>" & plugin_protocol & "</plugin_protocol>" & vbCrLf & _
        "<plugin_port>" & plugin_port & "</plugin_port>" & vbCrLf & _
        "<plugin_procedure_detection>" & plugin_procedure_detection & "</plugin_procedure_detection>" & vbCrLf & _
        "<plugin_procedure_exploit>" & plugin_procedure_exploit & "</plugin_procedure_exploit>" & vbCrLf & _
        "<plugin_detection_accuracy>" & plugin_detection_accuracy & "</plugin_detection_accuracy>" & vbCrLf & _
        "<plugin_exploit_accuracy>" & plugin_exploit_accuracy & "</plugin_exploit_accuracy>" & vbCrLf & _
        "<plugin_comment>" & plugin_comment & "</plugin_comment>" & vbCrLf
     
     'Add the bug data part 1
     PluginContent = PluginContent & _
        "<bug_published_name>" & bug_published_name & "</bug_published_name>" & vbCrLf & _
        "<bug_published_email>" & bug_published_email & "</bug_published_email>" & vbCrLf & _
        "<bug_published_web>" & bug_published_web & "</bug_published_web>" & vbCrLf & _
        "<bug_published_company>" & bug_published_company & "</bug_published_company>" & vbCrLf & _
        "<bug_published_date>" & bug_published_date & "</bug_published_date>" & vbCrLf & _
        "<bug_advisory>" & bug_advisory & "</bug_advisory>" & vbCrLf & _
        "<bug_produced_name>" & bug_produced_name & "</bug_produced_name>" & vbCrLf & _
        "<bug_produced_email>" & bug_produced_email & "</bug_produced_email>" & vbCrLf & _
        "<bug_produced_web>" & bug_produced_web & "</bug_produced_web>" & vbCrLf & _
        "<bug_affected>" & bug_affected & "</bug_affected>" & vbCrLf & _
        "<bug_not_affected>" & bug_not_affected & "</bug_not_affected>" & vbCrLf & _
        "<bug_false_positives>" & bug_false_positives & "</bug_false_positives>" & vbCrLf & _
        "<bug_false_negatives>" & bug_false_negatives & "</bug_false_negatives>" & vbCrLf & _
        "<bug_vulnerability_class>" & bug_vulnerability_class & "</bug_vulnerability_class>" & vbCrLf & _
        "<bug_description>" & bug_description & "</bug_description>" & vbCrLf & _
        "<bug_solution>" & bug_solution & "</bug_solution>" & vbCrLf & _
        "<bug_fixing_time>" & bug_fixing_time & "</bug_fixing_time>" & vbCrLf & _
        "<bug_exploit_availability>" & bug_exploit_availability & "</bug_exploit_availability>" & vbCrLf & _
        "<bug_exploit_url>" & bug_exploit_url & "</bug_exploit_url>" & vbCrLf & _
        "<bug_remote>" & bug_remote & "</bug_remote>" & vbCrLf & _
        "<bug_local>" & bug_local & "</bug_local>" & vbCrLf

    'Add the bug data part 2
    PluginContent = PluginContent & _
        "<bug_severity>" & bug_severity & "</bug_severity>" & vbCrLf & _
        "<bug_popularity>" & bug_popularity & "</bug_popularity>" & vbCrLf & _
        "<bug_simplicity>" & bug_simplicity & "</bug_simplicity>" & vbCrLf & _
        "<bug_impact>" & bug_impact & "</bug_impact>" & vbCrLf & _
        "<bug_risk>" & bug_risk & "</bug_risk>" & vbCrLf & _
        "<bug_nessus_risk>" & bug_nessus_risk & "</bug_nessus_risk>" & vbCrLf & _
        "<bug_iss_scanner_rating>" & bug_iss_scanner_rating & "</bug_iss_scanner_rating>" & vbCrLf & _
        "<bug_netrecon_rating>" & bug_netrecon_rating & "</bug_netrecon_rating>" & vbCrLf & _
        "<bug_check_tool>" & bug_check_tool & "</bug_check_tool>" & vbCrLf

    'Add the sources part 1
    PluginContent = PluginContent & _
        "<source_cve>" & source_cve & "</source_cve>" & vbCrLf & _
        "<source_certvu_id>" & source_certvu_id & "</source_certvu_id>" & vbCrLf & _
        "<source_cert_id>" & source_cert_id & "</source_cert_id>" & vbCrLf & _
        "<source_uscertta_id>" & source_uscertta_id & "</source_uscertta_id>" & vbCrLf & _
        "<source_securityfocus_bid>" & source_securityfocus_bid & "</source_securityfocus_bid>" & vbCrLf & _
        "<source_osvdb_id>" & source_osvdb_id & "</source_osvdb_id>" & vbCrLf & _
        "<source_secunia_id>" & source_secunia_id & "</source_secunia_id>" & vbCrLf & _
        "<source_securiteam_url>" & source_securiteam_url & "</source_securiteam_url>" & vbCrLf & _
        "<source_securitytracker_id>" & source_securitytracker_id & "</source_securitytracker_id>" & vbCrLf & _
        "<source_scip_id>" & source_scip_id & "</source_scip_id>" & vbCrLf & _
        "<source_tecchannel_id>" & source_tecchannel_id & "</source_tecchannel_id>" & vbCrLf & _
        "<source_heise_news>" & source_heise_news & "</source_heise_news>" & vbCrLf & _
        "<source_heise_security>" & source_heise_security & "</source_heise_security>" & vbCrLf & _
        "<source_aerasec_id>" & source_aerasec_id & "</source_aerasec_id>" & vbCrLf
     
    'Add the sources part 2
    PluginContent = PluginContent & _
        "<source_nessus_id>" & source_nessus_id & "</source_nessus_id>" & vbCrLf & _
        "<source_issxforce_id>" & source_issxforce_id & "</source_issxforce_id>" & vbCrLf & _
        "<source_snort_id>" & source_snort_id & "</source_snort_id>" & vbCrLf & _
        "<source_arachnids_id>" & source_arachnids_id & "</source_arachnids_id>" & vbCrLf & _
        "<source_mssb_id>" & source_mssb_id & "</source_mssb_id>" & vbCrLf & _
        "<source_mskb_id>" & source_mskb_id & "</source_mskb_id>" & vbCrLf & _
        "<source_netbsdsa_id>" & source_netbsdsa_id & "</source_netbsdsa_id>" & vbCrLf & _
        "<source_rhsa_id>" & source_rhsa_id & "</source_rhsa_id>" & vbCrLf & _
        "<source_ciac_id>" & source_ciac_id & "</source_ciac_id>" & vbCrLf & _
        "<source_literature>" & source_literature & "</source_literature>" & vbCrLf & _
        "<source_misc>" & source_misc & "</source_misc>"
        
    'Kill all useless lines to save space and increase performance
    PluginContentArray = Split(PluginContent, vbCrLf)
    PluginContentItemCount = UBound(PluginContentArray)
    PluginContent = vbNullString
    
    For i = 0 To PluginContentItemCount
        If InStr(4, PluginContentArray(i), "></") = 0 Then
                PluginContent = PluginContent & PluginContentArray(i) & vbCrLf
        End If
    Next i
        
    'Write the collected data into the file; the plugin name will be the file name
    Open Filename & ".plugin" For Output As 1
        Print #1, PluginContent
    Close

End Sub

Public Function HowManyLoadedPlugins() As Integer
    If frmMain.filATKPlugins.ListCount Then
        If frmMain.tvwPlugins.Nodes.Count Then
            On Error Resume Next 'Workaround to prevent kill if node is not available.
            HowManyLoadedPlugins = frmMain.tvwPlugins.Nodes("ATK ID").Children
        End If
    Else
        HowManyLoadedPlugins = 0
    End If
End Function

' ******************************************************************
' * This function extracts the possible ISBN number from a string. *
' * It works but there is one really nasty limitation:             *
' * 1. The ISBN numbers have to be written without the suggested   *
' *    delimiters as like spaces or dashes.                        *
' * This limitation should be fixed in an upcoming release.        *
' ******************************************************************

Public Function GetISBNFromString(TextString As String) As String
    Dim WordArray() As String
    Dim PossibleISBNNumber As String
    Dim i As Integer
    Dim j As Integer
    
    WordArray = Split(TextString, " ")
    
    For i = 0 To UBound(WordArray)
        'Reset the possible ISBN number for the next text block
        PossibleISBNNumber = vbNullString
        
        For j = 1 To Len(WordArray(i))
            If Len(PossibleISBNNumber) < 12 Then
                If Mid$(WordArray(i), j, 1) Like "[0-9]" Then
                    PossibleISBNNumber = PossibleISBNNumber & Mid$(WordArray(i), j, 1)
                ElseIf InStr(j, WordArray(i), "X") Then
                    PossibleISBNNumber = PossibleISBNNumber & Mid$(WordArray(i), j, 1)
                End If
                    
                If PossibleISBNNumber Like "#########?" Then
                    GetISBNFromString = PossibleISBNNumber
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Public Sub GenerateActualPluginList()
    Dim strFileContent As String
    Dim i As Integer
    Dim intLoadedPlugins As Integer

    'Set the progress bar to zero
    frmMain.SetProgress 0

    'Count the loaded plugins
    intLoadedPlugins = frmMain.filATKPlugins.ListCount - 1

    For i = 0 To intLoadedPlugins
        'Increase the progress bar. The On Error Resume Next prevents senseless
        'values that could lead to a programm error.
        On Error Resume Next
        frmMain.SetProgress (100 / intLoadedPlugins) * i
        
        'Everytime select the new plugin and do the check until finish
        'Set lsvPlugins.SelectedItem = lsvPlugins.ListItems(i)
        frmMain.filATKPlugins.ListIndex = i

        strFileContent = strFileContent & _
            plugin_id & ";" & frmMain.filATKPlugins.Filename & ";" & plugin_version & ";" & plugin_updated_date & ";" & plugin_filesize & vbCrLf
    
    Next i
    
    On Error Resume Next ' Needed if there are no write permissions
    Open PluginDirectory & "\pluginslist.txt" For Output As #1
        Print #1, strFileContent
    Close
    
    'Set the progress bar to 100
    frmMain.SetProgress 100
End Sub

Public Sub SetPluginSessionProcedure()
    If LenB(plugin_procedure_detection) Then
        session_procedure_type = "detection"
        session_procedure_commands = plugin_procedure_detection
    ElseIf LenB(plugin_procedure_exploit) Then
        session_procedure_type = "exploit"
        session_procedure_commands = plugin_procedure_exploit
    Else
        session_procedure_type = vbNullString
        session_procedure_commands = vbNullString
    End If
End Sub
