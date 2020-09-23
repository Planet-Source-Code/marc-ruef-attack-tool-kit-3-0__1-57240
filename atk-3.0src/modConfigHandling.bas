Attribute VB_Name = "modConfigHandling"
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 3.0 2004-10-11                                                           *
' * - Added a default value if it was not defined if logs should be activated.       *
' * - Added the whole procedures for handling the new logging security levels.       *
' ************************************************************************************

'For getting the username
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Change this "constant" on every new release to write the right software name
'and version.
Public Const SoftwareName As String = "Attack Tool Kit 3.0"
Public Const strProjectWebSiteURL As String = "http://www.computec.ch/projekte/atk/"

Public ActivateLogs As Boolean
Public ActivateSpeech As Boolean
Public ActivateSuggestions As Boolean
Public AlertingVulnFound As Boolean
Public AlertingVulnNotFound As Boolean
Public AttackMode As String
Public AttackTimeout As Long
Public DefaultSleepValue As Integer
Public DoICMPMapping As Boolean
Public DoNoDoSChecks As Boolean
Public DoSilentChecks As Boolean
Public HelpURL As String
Public LogsDirectory As String
Public LogsSecurityLevel As Integer
Public PluginDirectory As String
Public PluginsDownloadURL As String
'Public ReportsDirectory As String
Public ResponseDirectory As String
Public SuggestionsDirectory As String
Public SearchEngineURL As String
Public ScanIfICMPFails As Boolean
Public Target As String

Public application_configuration_filename As String
Public system_username As String

Public Sub LoadConfigFromFile(Optional ByRef strConfigurationFileName As String)
    Dim intFreeFile As Integer
    Dim TempString As String
    
    'This boolean values indicate that a value could be found. We need this state
    'to find missing or wrong input and correct them. This list is alphabetically until 1.1
    Dim ActivateLogsV As Boolean
    Dim ActivateSpeechV As Boolean
    Dim ActivateSuggestionsV As Boolean
    Dim AlertingVulnFoundV As Boolean
    Dim AlertingVulnNotFoundV As Boolean
    Dim AttackModeV As Boolean
    Dim AttackTimeoutV As Boolean
    Dim DefaultSleepValueV As Boolean
    Dim DoICMPMappingV As Boolean
    Dim DoNoDoSChecksV As Boolean
    Dim DoSilentChecksV As Boolean
    Dim HelpURLV As Boolean
    Dim LogsDirectoryV As Boolean
    Dim LogsSecurityLevelV As Boolean
    Dim PluginDirectoryV As Boolean
    Dim PluginsDownloadURLV As Boolean
    'Dim ReportsDirectoryV As Boolean
    Dim ScanIfICMPFailsV As Boolean
    Dim SearchEngineURLV As Boolean
    Dim SuggestionsDirectoryV As Boolean
    Dim TargetV As Boolean
        
    If LenB(strConfigurationFileName) Then
        application_configuration_filename = strConfigurationFileName
    Else
        application_configuration_filename = App.Path & "\configs\default.config"
    End If
    
    'WORKAROUND!
    ResponseDirectory = App.Path & "\responses\"
        
    'Check the existence of the config file
    If (Dir$(application_configuration_filename, 16) <> "") Then
        'Open and read the plugin file
        intFreeFile = FreeFile
        Open application_configuration_filename For Input As #intFreeFile
            Do While Not EOF(intFreeFile)
                Line Input #intFreeFile, TempString
                
                If Mid$(TempString, 1, 1) <> "#" Then
                    If InStr(1, TempString, "=", vbBinaryCompare) Then
                        If Mid$(TempString, 1, 16) = "PluginDirectory=" Then
                            PluginDirectory = Mid$(TempString, 17, Len(TempString))
                            If LenB(PluginDirectory) Then
                                PluginDirectoryV = True
                            End If
                        ElseIf Mid$(TempString, 1, 19) = "PluginsDownloadURL=" Then
                            PluginsDownloadURL = Mid$(TempString, 20, Len(TempString))
                            If LenB(PluginsDownloadURL) Then
                                PluginsDownloadURLV = True
                            End If
                        ElseIf Mid$(TempString, 1, 8) = "HelpURL=" Then
                            HelpURL = Mid$(TempString, 9, Len(TempString))
                            If LenB(HelpURL) Then
                                HelpURLV = True
                            End If
                        ElseIf Mid$(TempString, 1, 15) = "ActivateSpeech=" Then
                            ActivateSpeechV = True
                            If Mid$(TempString, 16, Len(TempString)) = 1 Then
                                ActivateSpeech = True
                            Else
                                ActivateSpeech = False
                            End If
                        ElseIf Mid$(TempString, 1, 20) = "ActivateSuggestions=" Then
                            ActivateSuggestionsV = True
                            If Mid$(TempString, 21, Len(TempString)) = 1 Then
                                ActivateSuggestions = True
                            Else
                                ActivateSuggestions = False
                            End If
                        ElseIf Mid$(TempString, 1, 18) = "AlertingVulnFound=" Then
                            AlertingVulnFoundV = True
                            If Mid$(TempString, 19, Len(TempString)) = 1 Then
                                AlertingVulnFound = True
                            Else
                                AlertingVulnFound = False
                            End If
                        ElseIf Mid$(TempString, 1, 21) = "AlertingVulnNotFound=" Then
                            AlertingVulnNotFoundV = True
                            If Mid$(TempString, 22, Len(TempString)) = 1 Then
                                AlertingVulnNotFound = True
                            Else
                                AlertingVulnNotFound = False
                            End If
                        ElseIf Mid$(TempString, 1, 21) = "SuggestionsDirectory=" Then
                            SuggestionsDirectory = Mid$(TempString, 22, Len(TempString))
                            If LenB(SuggestionsDirectory) Then
                                SuggestionsDirectoryV = True
                            End If
                        'ElseIf Mid$(TempString, 1, 17) = "ReportsDirectory=" Then
                        '    ReportsDirectoryV = True
                        '    ReportsDirectory = Mid$(TempString, 18, Len(TempString))
                        '    'Load another directory if it does not exists.
                        '    If Not (Dir$(ReportsDirectory, 16) <> "") Then
                        '        ReportsDirectory = App.Path
                        '    End If
                        ElseIf Mid$(TempString, 1, 14) = "AttackTimeout=" Then
                            AttackTimeoutV = True
                            AttackTimeout = Mid$(TempString, 15, Len(TempString))
                        ElseIf Mid$(TempString, 1, 18) = "DefaultSleepValue=" Then
                            DefaultSleepValueV = True
                            DefaultSleepValue = Mid$(TempString, 19, Len(TempString))
                        ElseIf Mid$(TempString, 1, 11) = "AttackMode=" Then
                            AttackModeV = True
                            AttackMode = Mid$(TempString, 12, Len(TempString))
                        ElseIf Mid$(TempString, 1, 15) = "DoSilentChecks=" Then
                            DoSilentChecksV = True
                            If Mid$(TempString, 16, Len(TempString)) = 1 Then
                                DoSilentChecks = True
                            Else
                                DoSilentChecks = False
                            End If
                        ElseIf Mid$(TempString, 1, 14) = "DoNoDoSChecks=" Then
                            DoNoDoSChecksV = True
                            If Mid$(TempString, 15, Len(TempString)) = 1 Then
                                DoNoDoSChecks = True
                            Else
                                DoNoDoSChecks = False
                            End If
                        ElseIf Mid$(TempString, 1, 14) = "DoICMPMapping=" Then
                            DoICMPMappingV = True
                            If Mid$(TempString, 15, Len(TempString)) = 1 Then
                                DoICMPMapping = True
                            Else
                                DoICMPMapping = False
                            End If
                        ElseIf Mid$(TempString, 1, 16) = "ScanIfICMPFails=" Then
                            ScanIfICMPFailsV = True
                            If Mid$(TempString, 17, Len(TempString)) = 1 Then
                                ScanIfICMPFails = True
                            Else
                                ScanIfICMPFails = False
                            End If
                        ElseIf Mid$(TempString, 1, 7) = "Target=" Then
                            Target = Mid$(TempString, 8, Len(TempString))
                            If LenB(Target) Then
                                TargetV = True
                            End If
                        ElseIf Mid$(TempString, 1, 13) = "ActivateLogs=" Then
                            ActivateLogsV = True
                            If Mid$(TempString, 14, Len(TempString)) = 1 Then
                                ActivateLogs = True
                            Else
                                ActivateLogs = False
                            End If
                        ElseIf Mid$(TempString, 1, 14) = "LogsDirectory=" Then
                            LogsDirectory = Mid$(TempString, 15, Len(TempString))
                            If LenB(LogsDirectory) Then
                                LogsDirectoryV = True
                            End If
                        ElseIf Mid$(TempString, 1, 18) = "LogsSecurityLevel=" Then
                            LogsSecurityLevelV = True
                            If Mid$(TempString, 19, Len(TempString)) = 0 Then
                                LogsSecurityLevel = 0
                            ElseIf Mid$(TempString, 19, Len(TempString)) = 1 Then
                                LogsSecurityLevel = 1
                            ElseIf Mid$(TempString, 19, Len(TempString)) = 2 Then
                                LogsSecurityLevel = 2
                            ElseIf Mid$(TempString, 19, Len(TempString)) = 3 Then
                                LogsSecurityLevel = 3
                            ElseIf Mid$(TempString, 19, Len(TempString)) = 4 Then
                                LogsSecurityLevel = 4
                            ElseIf Mid$(TempString, 19, Len(TempString)) = 5 Then
                                LogsSecurityLevel = 5
                            ElseIf Mid$(TempString, 19, Len(TempString)) = 6 Then
                                LogsSecurityLevel = 6
                            Else
                                LogsSecurityLevel = 7
                            End If
                        ElseIf Mid$(TempString, 1, 16) = "SearchEngineURL=" Then
                            SearchEngineURL = Mid$(TempString, 17, Len(TempString))
                            If LenB(SearchEngineURL) Then
                                SearchEngineURLV = True
                            End If
                        End If
                    End If
                End If
            Loop
        Close
    End If

    'Define default values if there is no config or no useful value in the config.
    'This is done to prevent false or missing input that would cause to an
    'undefined programm state.
    If PluginDirectoryV = False Then
        PluginDirectory = App.Path & "\plugins"
    End If
    
    If PluginsDownloadURLV = False Then
        PluginsDownloadURL = strProjectWebSiteURL & "plugins/pluginslist/"
    End If
    
    If HelpURLV = False Then
        HelpURL = strProjectWebSiteURL & "documentation/help/"
    End If
    
    If ActivateSuggestionsV = False Then
        ActivateSuggestions = True
    End If
        
    If ActivateSpeechV = False Then
        ActivateSpeech = False
    End If
            
    If AlertingVulnFoundV = False Then
        AlertingVulnFound = False
    End If
        
    If AlertingVulnNotFoundV = False Then
        AlertingVulnNotFound = False
    End If
    
    If ActivateLogsV = False Then
        ActivateLogs = True
    End If
        
    If SuggestionsDirectoryV = False Then
        SuggestionsDirectory = App.Path & "\suggestions"
    End If
        
    'If ReportsDirectoryV = False Then
    '    ReportsDirectory = App.Path & "\reports"
    'End If
        
    If LogsDirectoryV = False Then
        LogsDirectory = App.Path & "\logs"
    End If
        
    If LogsSecurityLevelV = False Then
        LogsSecurityLevel = 5
    End If
        
    If AttackTimeoutV = False Then
        AttackTimeout = 30000
    End If
        
    If DefaultSleepValueV = False Then
        DefaultSleepValue = 3000
    End If
        
    If AttackModeV = False Then
        AttackMode = "SingleCheck"
    End If
        
    If DoSilentChecksV = False Then
        DoSilentChecks = True
    End If
        
    If DoNoDoSChecksV = False Then
        DoNoDoSChecks = False
    End If
        
    If DoICMPMappingV = False Then
        DoICMPMapping = True
    End If
        
    If ScanIfICMPFailsV = False Then
        ScanIfICMPFails = False
    End If
        
    If TargetV = False Then
        Target = "127.0.0.1"
    End If
    
    If SearchEngineURLV = False Then
        SearchEngineURL = "http://www.google.com/search?q="
    End If
    
    'Change frame title so the user can see the next target
    frmMain.Caption = SoftwareName & " - " & Target
End Sub

Public Sub WriteConfigurationToFile(ByRef strConfigurationFileName As String)
    Dim intFreeFile As Integer
    Dim ConfigContent As String
    
    application_configuration_filename = strConfigurationFileName
    
    'Write the config file header
    ConfigContent = "#" & vbCrLf & _
        "# " & SoftwareName & " configuration file" & vbCrLf & _
        "# " & vbCrLf & _
        "#   Date       " & Date & vbCrLf & _
        "#   Time       " & Time & vbCrLf & _
        "#   File name  " & application_configuration_filename & vbCrLf & _
        "#   System     " & frmMain.wskTCPWinsock.Item(0).LocalIP & vbCrLf & _
        "#   User name  " & system_username & vbCrLf & _
        "#" & vbCrLf

    'Write a disclaimer
    ConfigContent = ConfigContent & _
        "# Disclaimer: This config file is generated automatically by the software" & vbCrLf & _
        "# itself during runtime. Please do not manually edit these values unless you" & vbCrLf & _
        "# do know what you're doing." & vbCrLf & _
        "#" & vbCrLf & _
        "# All values are shortly described. The left side specifies the variant were" & vbCrLf & _
        "# the data is saved and the right side defines the dynamicly saved value." & vbCrLf & _
        "# As it is used in most higher programming languages (e.g. Microsoft Visual" & vbCrLf & _
        "# Basic or ANSI C). The sharp sign can be used to uncomment a line. In this" & vbCrLf & _
        "# case the ATK uses the default value which is usually recommended." & vbCrLf & _
        "#" & vbCrLf & _
        "# See the online help, documentation and the official project web site" & vbCrLf & _
        "# http://www.computec.ch/projekte/atk/ for more details." & vbCrLf & _
        "#" & vbCrLf & vbCrLf
    
    'Write if logging should be done
    ConfigContent = ConfigContent & _
        "# The activate logs is a boolean variant were the activation for the logging" & vbCrLf & _
        "# feature is saved. The logging mechanism is used to do further analysis of" & vbCrLf & _
        "# scanning or debugging of the software. Activated logs may slow down the" & vbCrLf & _
        "# software a little bit. Activation of the logs is recommended. The value 0" & vbCrLf & _
        "# deactivates and 1 activates the logging. Activated logging with is the" & vbCrLf & _
        "# default value." & vbCrLf & _
        "#" & vbCrLf
    If ActivateLogs = True Then
        ConfigContent = ConfigContent & "ActivateLogs=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "ActivateLogs=0" & vbCrLf & vbCrLf
    End If
    
    'Write if speech output should be done
    ConfigContent = ConfigContent & _
        "# The activate speech is a boolean variant were the support for voice output" & vbCrLf & _
        "# is saved. This feature support the output of the application. Using spoken" & vbCrLf & _
        "# output slows down the software very much. The value 1 activates the speech" & vbCrLf & _
        "# and 0 deactivates it. The default value is 0 for deactivating the speech" & vbCrLf & _
        "# feature." & vbCrLf & _
        "#" & vbCrLf
    If ActivateSpeech = True Then
        ConfigContent = ConfigContent & "ActivateSpeech=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "ActivateSpeech=0" & vbCrLf & vbCrLf
    End If
    
    'Write the suggestions mode
    ConfigContent = ConfigContent & _
        "# The activate suggestion is a boolean value were the support for suggestions" & vbCrLf & _
        "# is saved. The value 1 stands for active and the opposit value 0 stands for" & vbCrLf & _
        "# not active. Activating the suggestions may slow down dedicated scans a" & vbCrLf & _
        "# little bit. But the suggestions are recommended for users who wants to be" & vbCrLf & _
        "# guided thru a penetration test." & vbCrLf & _
        "#" & vbCrLf
    If ActivateSuggestions = True Then
        ConfigContent = ConfigContent & "ActivateSuggestions=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "ActivateSuggestions=0" & vbCrLf & vbCrLf
    End If
    
    'Write if alerting if the bug is found should be done
    ConfigContent = ConfigContent & _
        "# The alerting vuln found is a boolean variant were the activation for" & vbCrLf & _
        "# messages if a vulnerability has been found is saved. This informs the user" & vbCrLf & _
        "# by a big message and may be useful in very long plugin testing attempts." & vbCrLf & _
        "# The value 1 activates the notification and 0 deactivates it. The default" & vbCrLf & _
        "# value is 0 for deactivated notification." & vbCrLf & _
        "#" & vbCrLf
    If AlertingVulnFound = True Then
        ConfigContent = ConfigContent & "AlertingVulnFound=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "AlertingVulnFound=0" & vbCrLf & vbCrLf
    End If
    
    'Write if alerting if the bug is found should be done
    ConfigContent = ConfigContent & _
        "# The alerting vuln not found is a boolean variant were the activation for" & vbCrLf & _
        "# messages if a vulnerability has not been found is saved. This informs the" & vbCrLf & _
        "# user by a big message and may be useful in very long plugin testing" & vbCrLf & _
        "# attempts. The value 1 activates the notification and 0 deactivates it. The" & vbCrLf & _
        "# default value is 0 for deactivated notification." & vbCrLf & _
        "#" & vbCrLf
    If AlertingVulnNotFound = True Then
        ConfigContent = ConfigContent & "AlertingVulnNotFound=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "AlertingVulnNotFound=0" & vbCrLf & vbCrLf
    End If
    
    'Write the attack mode
    ConfigContent = ConfigContent & _
        "# The attack mode is a string were the attack mode is saved. The possible" & vbCrLf & _
        "# values are SingleCheck and FullAudit. The first one is used to run" & vbCrLf & _
        "# singular plugins in a penetration test. The second one is used to run all" & vbCrLf & _
        "# loaded plugins in a security audit. The ATK was written to verify potential" & vbCrLf & _
        "# flaws and exploit verified vulnerabilities. So It is recommended to run in" & vbCrLf & _
        "# SingleCheck mode. The enumeration before exploiting with the ATK should be" & vbCrLf & _
        "# done with a vulnerability scanner as like Nessus. This because such are" & vbCrLf & _
        "# faster in security auditing than the ATK. The ATK is more an exploiting" & vbCrLf & _
        "# framework as like MetaSploit Framework or raccess." & vbCrLf & _
        "#" & vbCrLf
    If AttackMode = "SingleCheck" Then
        ConfigContent = ConfigContent & "AttackMode=SingleCheck" & vbCrLf & vbCrLf
    ElseIf AttackMode = "FullAudit" Then
        ConfigContent = ConfigContent & "AttackMode=FullAudit" & vbCrLf & vbCrLf
    End If
    
    'Write the attack timeout
    ConfigContent = ConfigContent & _
        "# The attack timeout is an integer value were the default value for timeouts" & vbCrLf & _
        "# during attacks is saved. The attack abords if it takes longer than this" & vbCrLf & _
        "# timeout value. This value is saved in milliseconds. This means 10000 stands" & vbCrLf & _
        "# for 10 seconds. Keep in mind that Microsoft Windows is not a real-time" & vbCrLf & _
        "# operating system as like QNX is. This is why very well defined values are" & vbCrLf & _
        "# not used as exactly as it may be wanted. Too short timeouts prevent a" & vbCrLf & _
        "# plugin to be successful and accurate. The recommended value is 30000 ms" & vbCrLf & _
        "# (this is 30 seconds timeout per attack)." & vbCrLf & _
        "#" & vbCrLf & _
        "AttackTimeout=" & AttackTimeout & vbCrLf & vbCrLf
    
    'Write the default sleep value
    ConfigContent = ConfigContent & _
        "# The default sleep value is an integer were the default sleep time for the" & vbCrLf & _
        "# sleep command is saved. This is used to let the application or a plugin" & vbCrLf & _
        "# wait a defined time value. This value is saved in milliseconds. This means" & vbCrLf & _
        "# 1000 stands for 1 second. Keep in mind that Microsoft Windows is not a" & vbCrLf & _
        "# real-time operating system as like QNX is. This is why very well defined" & vbCrLf & _
        "# values are not used as exactly as it may be wanted. Too short sleep values" & vbCrLf & _
        "# prevent a plugin to be successful and accurate. Too long sleep values will" & vbCrLf & _
        "# take a check longer to finish. The recommended value is 3000 ms (this is 3" & vbCrLf & _
        "# seconds timeout per attack)." & vbCrLf & _
        "#" & vbCrLf & _
        "DefaultSleepValue=" & DefaultSleepValue & vbCrLf & vbCrLf
    
    'Write if ICMP mapping should be done
    ConfigContent = ConfigContent & _
        "# The do icmp mapping is a boolean variant were the support for icmp/ping" & vbCrLf & _
        "# mapping is saved. Icmp mapping is used to determine the existence and" & vbCrLf & _
        "# reachability of a target before scanning. This may prevent attack attempts" & vbCrLf & _
        "# to non existing nor non reachable hosts. Such pre-verification may save" & vbCrLf & _
        "# time for further analysis. The mapping feature makes the software a bit" & vbCrLf & _
        "# slower and uses a bit more network ressources. But the feature is" & vbCrLf & _
        "# recommended to be more accurate. The value 1 activates the mapping" & vbCrLf & _
        "# feature and 0 deactivates it. The default value is 1 for activated icmp" & vbCrLf & _
        "# mapping." & vbCrLf & _
        "#" & vbCrLf
    If DoICMPMapping = True Then
        ConfigContent = ConfigContent & "DoICMPMapping=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "DoICMPMapping=0" & vbCrLf & vbCrLf
    End If
    
    'Write of denial of service checks should be done
    ConfigContent = ConfigContent & _
        "# The do no DoS checks is a boolean variant were the support for destructive" & vbCrLf & _
        "# and dangerous denial of service attacks is saved. You should deactivate the" & vbCrLf & _
        "# feature if the testing is done in a live environment that should not be" & vbCrLf & _
        "# harmed. The activation of DoS is recommended if the accuracy of a" & vbCrLf & _
        "# penetration test is very important. The value 1 activates the DoS save" & vbCrLf & _
        "# feature and 0 deactivates it. The default value is 0 for activated denial" & vbCrLf & _
        "# of service checks." & vbCrLf & _
        "#" & vbCrLf
    If DoNoDoSChecks = True Then
        ConfigContent = ConfigContent & "DoNoDoSChecks=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "DoNoDoSChecks=0" & vbCrLf & vbCrLf
    End If
    
    'Write if silent checks should be done
    ConfigContent = ConfigContent & _
        "# The silent check is a boolean variant were the support for silent checks is" & vbCrLf & _
        "# saved. Silent checks are working like the KB save feature in Nessus. The" & vbCrLf & _
        "# gathered data is used to verify other potential vulnerabilities without" & vbCrLf & _
        "# touching the target anymore. This makes the access attempts much faster and" & vbCrLf & _
        "# harder to detect by the target network. The silent check mode makes the" & vbCrLf & _
        "# software a bit slower. But the feature is recommended in all circumstances" & vbCrLf & _
        "# when very fast verification of potential vulnerabilities is required. The" & vbCrLf & _
        "# value 1 activates the silent check feature and 0 deactivates it. The" & vbCrLf & _
        "# default value is 1 for activated silent checks." & vbCrLf & _
        "#" & vbCrLf
    If DoSilentChecks = True Then
        ConfigContent = ConfigContent & "DoSilentChecks=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "DoSilentChecks=0" & vbCrLf & vbCrLf
    End If
    
    'Write the help URL
    ConfigContent = ConfigContent & _
        "# The help url is a string were the default url for access to the application" & vbCrLf & _
        "# online help is saved. This online help repository provides the information" & vbCrLf & _
        "# to handle the software. You are able to provide your own online help" & vbCrLf & _
        "# repository by putting an equivalent html based help on an accessible web" & vbCrLf & _
        "# server and specifying the online help repository url here. It is" & vbCrLf & _
        "# recommended to use the official online help repository at" & vbCrLf & _
        "# http://www.computec.ch/projekte/atk/documentation/help/" & vbCrLf & _
        "#" & vbCrLf & _
        "HelpURL=" & HelpURL & vbCrLf & vbCrLf
    
    'Write the Logs directory
    ConfigContent = ConfigContent & _
        "# The logs directory is as string were the default path name of the log files" & vbCrLf & _
        "# is saved. This data is relevant for the application to write and load the" & vbCrLf & _
        "# logging data. The default value is \logs" & vbCrLf & _
        "#" & vbCrLf & _
        "LogsDirectory=" & LogsDirectory & vbCrLf & vbCrLf
    
    'Write the logging security level
    ConfigContent = ConfigContent & _
        "# The log security level is an integer variant were the logging level is" & vbCrLf & _
        "# specified. This is the same as like the security level of syslog. Possible" & vbCrLf & _
        "# integer values range from 0 to 7. 7 are very important messages and 7 are" & vbCrLf & _
        "# for debugging only. As more messages are logged, as more ressources are" & vbCrLf & _
        "# used. Very verbose logging may slow down the software a bit. It is" & vbCrLf & _
        "# recommended to set the log level at 5 to get the most import messages." & vbCrLf & _
        "#" & vbCrLf
    If LogsSecurityLevel = 0 Then
        ConfigContent = ConfigContent & "LogsSecurityLevel=0" & vbCrLf & vbCrLf
    ElseIf LogsSecurityLevel = 1 Then
        ConfigContent = ConfigContent & "LogsSecurityLevel=1" & vbCrLf & vbCrLf
    ElseIf LogsSecurityLevel = 2 Then
        ConfigContent = ConfigContent & "LogsSecurityLevel=2" & vbCrLf & vbCrLf
    ElseIf LogsSecurityLevel = 3 Then
        ConfigContent = ConfigContent & "LogsSecurityLevel=3" & vbCrLf & vbCrLf
    ElseIf LogsSecurityLevel = 4 Then
        ConfigContent = ConfigContent & "LogsSecurityLevel=4" & vbCrLf & vbCrLf
    ElseIf LogsSecurityLevel = 5 Then
        ConfigContent = ConfigContent & "LogsSecurityLevel=5" & vbCrLf & vbCrLf
    ElseIf LogsSecurityLevel = 6 Then
        ConfigContent = ConfigContent & "LogsSecurityLevel=6" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "LogsSecurityLevel=7" & vbCrLf & vbCrLf
    End If
    
    'Write the plugin directory
    ConfigContent = ConfigContent & _
        "# The plugin directory is as string were the default path name of the plugins" & vbCrLf & _
        "# is saved This data is relevant for the application to load the checks and" & vbCrLf & _
        "# to run the access attempts. The default value is \plugins" & vbCrLf & _
        "#" & vbCrLf & _
        "PluginDirectory=" & PluginDirectory & vbCrLf & vbCrLf
    
    'Write the plugin download URL
    ConfigContent = ConfigContent & _
        "# The plugin download url is a string were the default url for access to the" & vbCrLf & _
        "# plugin repository is saved. This data is used to fetch the latest plugins" & vbCrLf & _
        "# and install them into the plugin directory. You are able to provide your" & vbCrLf & _
        "# own plugin repository server by putting your exported plugin list on an" & vbCrLf & _
        "# accessible web server and specifying the online plugin repository url here" & vbCrLf & _
        "# It is recommended to use the official plugin repository at" & vbCrLf & _
        "# http://www.computec.ch/projekte/atk/plugins/pluginslist/" & vbCrLf & _
        "#" & vbCrLf & _
        "PluginsDownloadURL=" & PluginsDownloadURL & vbCrLf & vbCrLf
    
    'Write if scan should be done if ICMP mapping fails
    ConfigContent = ConfigContent & _
        "# The scan if icmp mapping fails is a boolean variant were the support for" & vbCrLf & _
        "# scans if icmp mapping has been failed is saved. Icmp mapping is used to" & vbCrLf & _
        "# determine the existence and reachability of a target before scanning." & vbCrLf & _
        "# This may prevent attack attempts to non existing nor non reachable hosts." & vbCrLf & _
        "# Under some circumstances the target system does not not is able to react" & vbCrLf & _
        "# on icmp requests so the scan would fail. In this case this function should" & vbCrLf & _
        "# be activated to run a check if icmp mapping has been failed. The feature is" & vbCrLf & _
        "# recommended to be more accurate. The value 1 activates the overriding" & vbCrLf & _
        "# feature and 0 deactivates it. The default value is 0 for stopping scanning" & vbCrLf & _
        "# if icmp mapping fails." & vbCrLf & _
        "#" & vbCrLf
    If ScanIfICMPFails = True Then
        ConfigContent = ConfigContent & "ScanIfICMPFails=1" & vbCrLf & vbCrLf
    Else
        ConfigContent = ConfigContent & "ScanIfICMPFails=0" & vbCrLf & vbCrLf
    End If
    
    'Write the search engine URL
    ConfigContent = ConfigContent & _
        "# The search engine url is a string were the default query url for web" & vbCrLf & _
        "# searches. These are used for further ivestigation on the world wide web" & vbCrLf & _
        "# (e.g. looking for exploits). The software is providing a few well known" & vbCrLf & _
        "# search engine query urls. You are able to define your favorite search" & vbCrLf & _
        "# engine. It is important that the specified search engine allows queries as" & vbCrLf & _
        "# usual HTTP GET requests. In this case you are able to see your query string" & vbCrLf & _
        "# in the URL. Most web searches allow this method." & vbCrLf & _
        "#" & vbCrLf & _
        "SearchEngineURL=" & SearchEngineURL & vbCrLf & vbCrLf
    
    'Write the suggestionsdirectory
    ConfigContent = ConfigContent & _
        "# The suggestions directory is a string were the default directory for all" & vbCrLf & _
        "# the suggestions is saved. This suggestions repository helps new users to" & vbCrLf & _
        "# define the next steps after running an audit attempt. The default" & vbCrLf & _
        "# directory is \suggestions" & vbCrLf & _
        "#" & vbCrLf & _
        "SuggestionsDirectory=" & SuggestionsDirectory & vbCrLf & vbCrLf
    
    'Write the Target
    ConfigContent = ConfigContent & _
        "# The target is a string were the target for the checking is specified. In" & vbCrLf & _
        "# here host names and ip addresses may be defined. Do not scan ressources" & vbCrLf & _
        "# without premission of the owner of the administrator. The default value is" & vbCrLf & _
        "# the loopback ip address 127.0.0.1 to allow check of the own localhost." & vbCrLf & _
        "#" & vbCrLf & _
        "Target=" & Target & vbCrLf & vbCrLf
    
    'Write the config in the config gile
    On Error Resume Next ' Needed if there are no write permissions
    intFreeFile = FreeFile
    Open strConfigurationFileName For Output As #intFreeFile
        Print #intFreeFile, ConfigContent
    Close
    
    'Show the frame title with configuration file name
    frmConfiguration.Caption = "Configuration - " & application_configuration_filename
    
    'Change frame title so the user can see the next target
    frmMain.Caption = SoftwareName & " - " & Target
End Sub

Public Sub LoadUserName()
    Dim strTemp As String
    
    strTemp = String(255, 0)
    GetUserName strTemp, 255
    system_username = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
End Sub
