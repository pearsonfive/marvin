Imports OTIS

Public Class TopicWatcher : Inherits WatcherBase

    Private _messagesFoundUpToLastTime As Hashtable
    Private _combinationLogic As String
    Private _watcherInfo As WatcherInfo

    Private WithEvents _gunTopic As OTIS.GunTopic


    ''' <summary>
    ''' Constructor for QueueWatcher
    ''' </summary>
    ''' <param name="ThisJob">The job object that we are going to watch for</param>
    ''' <remarks>SDCP 01-Nov-2006</remarks>
    Sub New(ByVal ThisJob As Job)

        '--- store the active trigger
        _activeJob = ThisJob

        '--- instantiate the hashtables
        _messagesFoundUpToLastTime = New Hashtable

        '--- don't want to stop yet
        _stopWatcher = False

        '--- load holidays if required
        LoadHolidays()

        '--- extract setup info from the job definition
        ExtractSetupInfo()

        '--- start polling
        StartPolling()

    End Sub


    ''' <summary>
    ''' extract setup info from the job definition 
    ''' </summary>
    ''' <remarks>SDCP 10-Oct-2006</remarks>
    Protected Overrides Sub ExtractSetupInfo()

        '--- base watcher settings:

        '--- poll interval
        '        _pollInterval = CInt(_activeJob.Trigger.TriggerInfoXml.SelectSingleNode("watcherbase/pollinterval").InnerText)
        _pollFrequencyUnit = _activeJob.Trigger.TriggerInfoXml.SelectSingleNode("watcherbase/pollfrequency/@unit").InnerText
        _pollFrequencyAmount = CInt(_activeJob.Trigger.TriggerInfoXml.SelectSingleNode("watcherbase/pollfrequency/@amount").InnerText)

        '--- get the combination logic
        Try
            _combinationLogic = _activeJob.Trigger.TriggerInfoXml.SelectSingleNode("watcherbase/@logic").Value
        Catch
            _combinationLogic = "and"
        End Try

        '--- de-serialise the info from the xml
        Dim ndWatcher As Xml.XmlNode = _activeJob.Trigger.TriggerInfoXml.SelectSingleNode("descendant::watcher[1]")

        '--- create a SetupInfo for the watcher
        _watcherInfo = New WatcherInfo

        With ndWatcher
            _watcherInfo.EMSServerName = .SelectSingleNode("emsserver").InnerText
            _watcherInfo.EMSUsername = .SelectSingleNode("emsusername").InnerText
            _watcherInfo.EMSPassword = .SelectSingleNode("emspassword").InnerText
            _watcherInfo.TopicName = .SelectSingleNode("topic").InnerText

            'xTODO: add MessageSelector to TopicWatcher
            'aWatcherInfo.MessageSelector = .SelectSingleNode("messageselector").InnerText

            '--- extra output fields
            _watcherInfo.ExtraOutputFields = New Hashtable
            For Each ndExtra As Xml.XmlNode In ndWatcher.SelectNodes("extraoutputfields/*")
                _watcherInfo.ExtraOutputFields.Add(ndExtra.Name, ndExtra.InnerXml)
            Next

        End With

    End Sub




    ''' <summary>
    ''' one off setup of the Topic subscriber as a listener
    ''' </summary>
    ''' <remarks>SDCP 01-Nov-2006</remarks>
    Protected Overrides Sub Poller()


        With _watcherInfo

            Try

                Dim anOTIS As OTIS.OTIS = New OTIS.OTIS(.EMSServerName, .EMSUsername, .EMSPassword)

                '--- get a topic object
                If _gunTopic Is Nothing Then
                    _gunTopic = anOTIS.CreateTopic
                End If

                '--- subscribe to the queue
                Call _gunTopic.SubscribeToTopic(.EMSServerName, .TopicName, .EMSUsername, .EMSPassword)

            Catch ex As Exception
                Stop
            End Try

        End With

    End Sub



    ''' <summary>
    ''' helper sub-class to store settings for each watcher
    ''' </summary>
    ''' <remarks></remarks>
    Class WatcherInfo
        Public EMSServerName As String
        Public EMSUsername As String
        Public EMSPassword As String
        Public TopicName As String
        Public MessageSelector As String
        Public ExtraOutputFields As Hashtable
    End Class


    ''' <summary>
    ''' when a message is received then need to deserialise the message contents
    ''' and write a new message to the Triggers queue
    ''' </summary>
    ''' <param name="TopicMessage"></param>
    ''' <remarks></remarks>
    Private Sub BridgeToQueue(ByVal TopicMessage As OTIS.GunMessage)

        Dim xmlDoc As Xml.XmlDocument = New Xml.XmlDocument
        Dim htNameValuePairs As Hashtable = New Hashtable

        Select Case TopicMessage.MessageType

            Case "TextMessage"

                '--- load the message text as XML
                xmlDoc.LoadXml(TopicMessage.MessageText)

                '--- output each name-value pair to the hashtable
                For Each xNode As Xml.XmlNode In xmlDoc.DocumentElement.SelectNodes("child::*")

                    htNameValuePairs.Add(xNode.Name, xNode.InnerXml)

                Next

            Case "MapMessage"
                '
                '


        End Select

        With htNameValuePairs

            '--- add some TopicWatcher specific fields
            .Add("originalemsserver", _watcherInfo.EMSServerName)
            .Add("originaltopic", _watcherInfo.TopicName)
            .Add("originaljmsmessageid", TopicMessage.MessageID)

            '--- any extra output fields
            For Each de As DictionaryEntry In _watcherInfo.ExtraOutputFields
                .Add(de.Key.ToString, de.Value.ToString)
            Next

        End With

        '--- post to Triggers queue
        MyMessenger.PostMessage(MyMessenger.Config.TriggersQ, "", "", "", htNameValuePairs)


    End Sub


    ''' <summary>
    ''' handler for a new topic message
    ''' </summary>
    ''' <param name="MessageReceived"></param>
    ''' <remarks></remarks>
    Private Sub _gunTopic_MessageReceived(ByVal MessageReceived As OTIS.GunMessage) Handles _gunTopic.MessageReceived

        BridgeToQueue(MessageReceived)

    End Sub



End Class

