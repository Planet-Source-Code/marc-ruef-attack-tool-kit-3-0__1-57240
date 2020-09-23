Attribute VB_Name = "modErrorHandling"
Option Explicit

Public Sub errPluginsDirectoryEmpty()
    WriteLogEntry "In " & PluginDirectory & " no plugins could be found.", 5
    
    If MsgBox("No plugins could be loaded because the default plugin directory" & vbCrLf & _
        PluginDirectory & vbCrLf & _
        "is empty! No predefined checks are possible at the moment." & vbCrLf & _
        "Please check the plugins directory configuration." & vbCrLf & vbCrLf & _
        "Would you like to start the AutoUpdate to download the latest ATK plugins?", _
        vbYesNo + vbInformation, "Attack Tool Kit load plugins error") = vbYes Then
        
        'Open the AutoUpdate
        Call frmMain.mnuPluginsDownloadTheLatestPluginsItem_Click
    Else
        WriteLogEntry "Opening AutoUpdate to get the latest plugins has been manually aborded.", 4
    End If
End Sub

Public Sub errPluginsDirectoryNotExist()
    'Error message if the plugin directory does not exists
    WriteLogEntry "The plugin directory " & PluginDirectory & " does not exists.", 3
    
    If MsgBox("No plugins could be loaded because the default plugin directory" & vbCrLf & _
        PluginDirectory & vbCrLf & _
        "does not exists! No predefined checks are possible at the moment." & vbCrLf & _
        "Please check the plugins directory configuration." & vbCrLf & vbCrLf & _
        "Would you like to create the plugins directory " & vbCrLf & _
        PluginDirectory & "?", vbYesNo + vbInformation, "Attack Tool Kit load plugins error") = vbYes Then
        
        'Make the plugin directory
        On Error Resume Next
        MkDir (PluginDirectory)
        WriteLogEntry "Plugins directory " & PluginDirectory & " created.", 6
    Else
        WriteLogEntry "Creating the plugin directory " & PluginDirectory & _
            " has been manually aborded.", 3
    End If
End Sub

Public Sub errLogDirectoryNotExist()
    'Developer note: We cannot use the logging feature in this procedure because we would
    'get a nasty recursive routine without an exit.
    
    If application_log_directory = False Then
        If MsgBox("No file logging could be done because the default logs directory" & vbCrLf & _
            LogsDirectory & vbCrLf & _
            "does not exists! No additionall debugging was possible until now." & vbCrLf & vbCrLf & _
            "Would you like to create the logs directory " & vbCrLf & _
            LogsDirectory & "?", vbYesNo + vbInformation, "Attack Tool Kit precheck logs warning") = vbYes Then
        
            'Make the logs directory
            On Error Resume Next 'Skip the mkdir command if there are no write permissions
            MkDir (LogsDirectory)
            WriteLogEntry "Logs directory " & LogsDirectory & " created.", 6
        Else
            'Set the value that no log directory is wished. All further error messages
            'in this field will be ignored and not shown.
            application_log_directory = True
        End If
    End If
End Sub

Public Sub errLogDirectoryEmpty()
    WriteLogEntry "In " & LogsDirectory & " no log files could be found.", 4
    
    If MsgBox("No log files could be found because the default log directory" & vbCrLf & _
        LogsDirectory & vbCrLf & _
        "is empty! No further application analysis possible at the moment." & vbCrLf & _
        "Please check the log directory configuration." & vbCrLf & vbCrLf & _
        "Would you like to load a specific log file to start a log analysis?", _
        vbYesNo + vbInformation, "Attack Tool Kit load log error") = vbYes Then
        
        'Open the AutoUpdate
        Call frmLog.mnuFileOpenItem_Click
    Else
        WriteLogEntry "Opening a specific log file has been manually aborded.", 4
    End If
End Sub

Public Sub errSuggestionsDirectoryNotExist()
    'Error message if the plugin directory does not exists
    WriteLogEntry "The suggestions directory " & SuggestionsDirectory & " does not exist.", 3
    
    If MsgBox("No suggestions could be loaded because the default suggestions directory" & vbCrLf & _
        SuggestionsDirectory & vbCrLf & _
        "does not exists! No additionall suggestions are possible at the moment." & vbCrLf & _
        "Please check the suggestions directory configuration." & vbCrLf & vbCrLf & _
        "Would you like to create the suggestions directory " & vbCrLf & _
        SuggestionsDirectory & "?", vbYesNo + vbInformation, "Attack Tool Kit suggestions error") = vbYes Then
            
        'Make the suggestions directory
        On Error Resume Next 'Skip the mkdir command if there are no write permissions
        MkDir (SuggestionsDirectory)
        WriteLogEntry "Suggestions directory " & SuggestionsDirectory & " created.", 6
    Else
        WriteLogEntry "Creating the suggestions directory " & SuggestionsDirectory & _
            " has been manually aborded.", 4
    End If
End Sub

'Public Sub errReportDirectoryNotExist()
'    'Error message if the plugin directory does not exists
'    WriteLogEntry "The reports directory " & ReportSdirectory & " does not exist.", 3
'
'    If MsgBox("No reports could be cached because the default reports directory" & vbCrLf & _
'        ReportSdirectory & vbCrLf & _
'        "does not exists! No further analysis was possible until now." & vbCrLf & vbCrLf & _
'        "Would you like to create the suggestions directory " & vbCrLf & _
'        ReportsDirectory & "?", vbYesNo + vbInformation, "Attack Tool Kit report warning") = vbYes Then
'
'        'Make the suggestions directory
'        On Error Resume Next 'Skip the mkdir command if there are no write permissions
'        MkDir (ReportsDirectory)
'    End If
'End Sub

Public Sub errPluginDoesNotExist(ByRef strPluginFileName As String)
    WriteLogEntry "The plugin " & strPluginFileName & " does not exist anymore.", 2
    
    If MsgBox("The specified plugin " & strPluginFileName & vbCrLf & _
        "does not exist anymore. It may be deleted since the last access. You are not able to use the plugin at the moment." & vbCrLf & vbCrLf & _
        "Please check the plugins directory configuration or run the AutoUpdate to download the latest plugins." & vbCrLf & vbCrLf & _
        "Would you like to start the AutoUpdate to re-initialize your local ATK plugins repository?", _
        vbYesNo + vbInformation, "Attack Tool Kit load plugin error") = vbYes Then
        
        'Open the AutoUpdate
        Call frmMain.mnuPluginsDownloadTheLatestPluginsItem_Click
    Else
        WriteLogEntry "Opening AutoUpdate to get the latest plugins as been manually aborded.", 4
    End If
End Sub

Public Sub errPluginDataMissing(ByRef strMissingDataName As String, ByRef strPluginFileName As String, ByRef intPluginID As String)
    
    'Write a log entry about the error
    WriteLogEntry "Important attack data " & strMissingDataName & " is missing. Check aborded.", 1
    
    'Show the error message
    If MsgBox("Important attack data " & strMissingDataName & " is missing." & vbCrLf & vbCrLf & _
        "You will not be able to run the plugin " & intPluginID & vbCrLf & _
        " (" & strPluginFileName & ") correctly." & vbCrLf & vbCrLf & _
        "Would you like to open the Attack Editor to check the error manually?", _
        vbYesNo + vbInformation, "Attack Tool Kit plugin data error") = vbYes Then
    
        'Show the attack editor to eliminate the check error
        frmAttackEditor.Visible = True
    Else
        WriteLogEntry "Opening the Attack Editor to check the missing data manually has been manually aborded.", 3
    End If
End Sub

Public Sub errTargetWrongSpecification()
    'Error message if the has been specified in a wrong way
    WriteLogEntry "The target has been specified in a wrong way.", 4
    
    MsgBox "You have specified the target in a wrong way that is not supported by this version" & vbCrLf & _
        "of the Attack Tool Kit (ATK)." & vbCrLf & vbCrLf & _
        "You can specify host names (e.g. www.computec.ch) or IP addresses (e.g." & vbCrLf & _
        "192.168.0.1) only. Your input has been re-written to prevent run-time errors." & vbCrLf & vbCrLf & _
        "Please check the new target definition to get the wanted match for your" & vbCrLf & _
        "attack.", vbOKOnly + vbInformation, "Attack Tool Kit target error"
End Sub
