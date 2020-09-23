Attribute VB_Name = "modResponseHandling"
Option Explicit

Public LastResponse As String
Public LastResponseTime As String

Public Sub ClearAllResponseVariables()
    LastResponseTime = vbNullString
    LastResponse = vbNullString
End Sub

Public Sub WriteLastResponseToFile()
    
    If Not (Dir$(ResponseDirectory, 16) <> "") Then
        On Error Resume Next    'Prevent errors if the device is write protected
        MkDir (ResponseDirectory)
    End If
    
    Open ResponseDirectory & Target & "-" & plugin_port & ".txt" For Output As #1
        On Error Resume Next    'Prevent errors if the device is write protected
        Print #1, LastResponse
    Close
End Sub

Public Sub LoadLatestResponse()
    If IsFormVisible("frmAttackResponse") = True Then
        frmAttackResponse.PrepareTabs
        
        frmAttackResponse.lblHost.Caption = Target
        frmAttackResponse.lblPort.Caption = plugin_port
        frmAttackResponse.lblTime.Caption = LastResponseTime
        frmAttackResponse.txtLastResponse.Text = LastResponse
        frmAttackResponse.lblLength.Caption = Len(LastResponse) & _
            " bytes"
    End If
End Sub
