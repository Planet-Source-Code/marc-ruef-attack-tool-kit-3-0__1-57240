Attribute VB_Name = "modBrowserCall"
Option Explicit

'Declare the function for the browser call
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
    lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Sub OpenProjectWebsite()
    'Load the project web site
    Call ShellExecute(frmMain.hwnd, "Open", strProjectWebSiteURL, "", App.Path, 1)
End Sub

Public Sub OpenOnlineHelp(Optional ByRef strSubDirectory As String)
    Dim strFullOnlineHelpURL As String
    
    strFullOnlineHelpURL = HelpURL & strSubDirectory
    
    If LenB(strFullOnlineHelpURL) = 0 Then
        strFullOnlineHelpURL = strProjectWebSiteURL
    End If
    
    'Load the online help
    WriteLogEntry "Opening the online help URL " & strFullOnlineHelpURL & " ...", 6
    Call ShellExecute(frmMain.hwnd, "Open", strFullOnlineHelpURL, "", App.Path, 1)
End Sub

Public Sub OpenOnlineSearch(Optional ByRef strSearchString As String)
    Dim strFullSearchURL As String
    
    strFullSearchURL = SearchEngineURL & strSearchString
    
    If LenB(strFullSearchURL) = 0 Then
        strFullSearchURL = "http://www.google.com"
    End If
    
    'Load the online search
    WriteLogEntry "Opening the search engine URL " & strFullSearchURL & " ...", 6
    Call ShellExecute(frmMain.hwnd, "Open", strFullSearchURL, "", App.Path, 1)
End Sub
