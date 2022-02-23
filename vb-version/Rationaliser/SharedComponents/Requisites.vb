Imports System.IO

Module Requisites

    ''' <summary>
    ''' Set up the messaging, and hence the config classes
    ''' </summary>
    ''' <param name="sConfigFile"></param>
    ''' <returns></returns>
    ''' <remarks>Martin Leyland, November 2006</remarks>
    ''' 

    Public Function SetUpForMessaging(ByVal sConfigFile As String) As Messenger

        Dim sMessage As String = "", i As Integer
        Try
            MyMessenger = New Messenger(sConfigFile, sTHIS_APP)
            If MyMessenger.Config.Errors.Count > 0 Then
                For i = 1 To MyMessenger.Config.Errors.Count
                    sMessage = sMessage & MyMessenger.Config.Errors(i).ToString & vbCr
                Next
            Else
                Return MyMessenger
                Exit Function
            End If

        Catch ex As Exception
            sMessage = ex.Message & vbCr
            sMessage = sMessage & ex.StackTrace & vbCr

        End Try

        'Check that the all loaded properly
        If sMessage <> "" Then
            LastDitchLog(sMessage)
        End If
        Return Nothing

    End Function

    ''' <summary>
    ''' Get the config file from the cmd line params
    ''' </summary>
    ''' <returns>Full path to the config file</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function GetConfigFileName() As String

        Dim sConfigFile As String = ""

        sConfigFile = MySwitches("config")
       
        If sConfigFile = "" Then
            sConfigFile = sDEFAULT_CONFIG_FILE
        End If

        Return sConfigFile

    End Function

    ''' <summary>
    ''' Something gone wrong, we have no messaging, so log to file
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <returns></returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>

    Public Function LastDitchLog(ByVal Message As String) As Boolean


        Dim objReader As StreamWriter
        Try
            objReader = New StreamWriter(sNETWORK_ERROR_LOGFILE & Format(Now, "yyyyMMdd HHmmss") & "_ConfigError.log")
        Catch ex As Exception

            '--- should be "docs and settings"
            objReader = New StreamWriter(sLOCAL_ERROR_LOGFILE & Format(Now, "yyyyMMdd HHmmss") & "_ConfigError.log")

        End Try
        objReader.WriteLine(Message)

        Try
            objReader.WriteLine(HarvestProcessingDetails("commandline|switches|config"))
        Catch
        Finally
            objReader.Close()
        End Try

        '--- that's all folks!
        End

    End Function


    Public Function HarvestProcessingDetails(ByVal What As String) As String
        Dim sDetails As String = ""

        On Error Resume Next

        sDetails += vbCrLf
        sDetails += vbCrLf

        If InStr(What, "commandline") <> 0 Then

            sDetails += "" & vbCrLf
            sDetails += "COMMAND LINE" & vbCrLf
            sDetails += "============" & vbCrLf
            sDetails += Environment.CommandLine & vbCrLf

        End If

        If InStr(What, "switches") <> 0 Then

            sDetails += "" & vbCrLf
            sDetails += "SWITCHES" & vbCrLf
            sDetails += "========" & vbCrLf
            For Each de As DictionaryEntry In MySwitches._switchesHashTable
                sDetails += de.Key.ToString & " = " & de.Value.ToString & vbCrLf
            Next

        End If

        If InStr(What, "config") <> 0 Then

            sDetails += "" & vbCrLf
            sDetails += "CONFIG FILE" & vbCrLf
            sDetails += "========" & vbCrLf
            sDetails += MyMessenger.Config.DomDoc.OuterXml & vbCrLf

        End If

        Return sDetails

    End Function



End Module

