Module ActionModule


    '**********************************************************
    '* Controller for processing
    '* 
    '*
    '*
    '* SDCP 13-Sep-2006
    '**********************************************************
    Public Sub StartProcessing()

        '--- load the trigger info
        LoadJobDefinitions()

        '--- start polling
        StartPolling()

    End Sub


    '**********************************************************
    '* load the trigger info XML and parse as appropriate
    '* 
    '*
    '*
    '* SDCP 25-Sep-2006
    '**********************************************************
    Private Sub LoadJobDefinitions()


        Dim jobDefinitions As New Xml.XmlDocument

        Try

            jobDefinitions.Load(MySwitches("jobdefinitions"))

        Catch ex As Xml.XmlException
            Try
                MyMessenger.AtomInfo += "<error>Failed to parse job def file</error>"
                MyMessenger.AtomInfo += "<errormessage>" & ex.Message & "</errormessage>"
                MyMessenger.AtomInfo += "<jobdefaddress>" & CDataBlock(MySwitches("jobdefinitions")) & "</jobdefaddress>"
                MyMessenger.PostMessage(MyMessenger.Config.StatusQ, "", "", "", GetWatcherErrorNVPs)
                MyMessenger.SendEmail("DL-RatesDeskDev@ubs.com", "DL-RatesDeskDev@ubs.com", "Watcher failure on " & My.Computer.Name, "Failed to parse [include] file" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace & HarvestProcessingDetails("commandline|switches"))

            Catch
            End Try

            End

        End Try

        parser = New JobDefinitionParser
        With parser
            .JobDefinitionsXml = jobDefinitions
            .Parse()
        End With

        '--- get the triggers collection
        '  TriggerCollection = parser.Triggers

    End Sub


    '**********************************************************
    '* start the polling process
    '* 
    '*
    '*
    '* SDCP 26-Sep-2006
    '**********************************************************
    Private Sub StartPolling()
        Manager = New WatcherManager
        Manager.CreateWatcherThreads()
    End Sub

End Module
