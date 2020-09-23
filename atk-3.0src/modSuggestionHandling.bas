Attribute VB_Name = "modSuggestionHandling"
Option Explicit

Public suggestion_name As String
Public suggestion_trigger As String
Public suggestion_description As String
Public suggestion_todo As String

Public Sub ReadSuggestionFromFile(Filename As String)
    Dim Temp As String              'The temporary file output
    Dim SuggestionContent As String 'The plugin content itself
    
    Dim TempArray() As String       'A temporary array for the splitting and parsing
   
    'Open and read the plugin file
    Open App.Path & "\suggestions\" & Filename For Input As 1
        Do While Not EOF(1)
            Line Input #1, Temp
                SuggestionContent = SuggestionContent & Temp
        Loop
    Close
        
    'Get the data fields and write them into the public variables
    TempArray = Split(SuggestionContent, "<name>")
    TempArray = Split(TempArray(1), "</name>")
    If Len(TempArray(0)) <> Len(SuggestionContent) Then
        suggestion_name = TempArray(0)
    Else
        suggestion_name = vbNullString
    End If
        
    'Get the data fields and write them into the public variables
    TempArray = Split(SuggestionContent, "<trigger>")
    TempArray = Split(TempArray(1), "</trigger>")
    If Len(TempArray(0)) <> Len(SuggestionContent) Then
        suggestion_trigger = TempArray(0)
    Else
        suggestion_trigger = vbNullString
    End If
    
    'Get the data fields and write them into the public variables
    TempArray = Split(SuggestionContent, "<description>")
    TempArray = Split(TempArray(1), "</description>")
    If Len(TempArray(0)) <> Len(SuggestionContent) Then
        suggestion_description = TempArray(0)
    Else
        suggestion_description = vbNullString
    End If
    
    'Get the data fields and write them into the public variables
    TempArray = Split(SuggestionContent, "<suggestion>")
    TempArray = Split(TempArray(1), "</suggestion>")
    If Len(TempArray(0)) <> Len(SuggestionContent) Then
        suggestion_todo = TempArray(0)
    Else
        suggestion_todo = vbNullString
    End If
    
End Sub
