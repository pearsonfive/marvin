Imports OTIS

Public Class PollingWatcher : Inherits WatcherBase

    '    Private _queueName As String
    '    Private _messageSelector As String

    Private _combinationLogic As String
    Private _watcherInfoList As List(Of WatcherInfo)



    ''' <summary>
    ''' Constructor for QueueWatcher
    ''' </summary>
    ''' <param name="ThisJob">The job object that we are going to watch for</param>
    ''' <remarks>SDCP 01-Nov-2006</remarks>
    Sub New(ByVal ThisJob As Job)

        '--- store the active trigger
        _activeJob = ThisJob

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

        '--- prep the list
        _watcherInfoList = New List(Of WatcherInfo)

        '--- base watcher settings:

        '--- poll interval
        _pollFrequencyUnit = _activeJob.Trigger.TriggerInfoXml.SelectSingleNode("watcherbase/pollfrequency/@unit").InnerText
        _pollFrequencyAmount = CInt(_activeJob.Trigger.TriggerInfoXml.SelectSingleNode("watcherbase/pollfrequency/@amount").InnerText)



        '--- get the combination logic
        Try
            _combinationLogic = _activeJob.Trigger.TriggerInfoXml.SelectSingleNode("watcherbase/@logic").Value
        Catch
            _combinationLogic = "and"
        End Try

        '--- de-serialise the info from the xml
        For Each ndWatcher As Xml.XmlNode In _activeJob.Trigger.TriggerInfoXml.SelectNodes("descendant::watcher")

            '--- create a SetupInfo for each watcher
            Dim aWatcherInfo As New WatcherInfo

            With ndWatcher
                'aWatcherInfo.PollAtStartup = CBool(.SelectSingleNode("pollatstartup").InnerText)
            End With

            '--- add to the list
            _watcherInfoList.Add(aWatcherInfo)

        Next

    End Sub




    ''' <summary>
    ''' polls the queue for changes since last poll time
    ''' </summary>
    ''' <remarks>SDCP 01-Nov-2006</remarks>
    Protected Overrides Sub Poller()

        '--- poll for each watcher
        For Each aWatcherInfo As WatcherInfo In _watcherInfoList

            With aWatcherInfo

                '--- create a hashtable for the found triggers
                .TriggersFound = New Hashtable
                Dim fti As New FoundTriggerInfo
                fti.PulledTrigger = "doesn't matter"
                .TriggersFound.Add("this", fti)
                
            End With

        Next 'aWatcherInfo


            '--- (assume found)
            FoundValidTrigger = True

            '--- all watcher info objects have to have trigger files
            For Each aWatcherInfo As WatcherInfo In _watcherInfoList

                '--- if there is a watcher with no found files then flag is false and can exit
                If aWatcherInfo.TriggersFound.Count = 0 Then
                    FoundValidTrigger = False
                    Exit For
                End If

            Next



        If FoundValidTrigger = True Then

            '--- populate the TriggersFound table
            _triggersFound = New Hashtable
            For Each aWatcherInfo As WatcherInfo In _watcherInfoList

                For Each de As DictionaryEntry In aWatcherInfo.TriggersFound

                    _triggersFound.Add(de.Key, de.Value)

                Next

            Next aWatcherInfo

            CreateMessages()
        End If

    End Sub



    ''' <summary>
    ''' helper sub-class to store settings for each watcher
    ''' </summary>
    ''' <remarks></remarks>
    Class WatcherInfo : Inherits WatcherBase.WatcherInfoBase
        '        Public PollAtStartup As Boolean
    End Class


End Class
