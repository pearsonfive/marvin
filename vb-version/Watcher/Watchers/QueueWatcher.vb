Imports OTIS
Imports System.Xml


Public Class QueueWatcher : Inherits WatcherBase

    '    Private _queueName As String
    '    Private _messageSelector As String

    Private _messagesFoundUpToLastTime As Hashtable
    Private _useCombination As Boolean
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

        '--- prep the list
        _watcherInfoList = New List(Of WatcherInfo)

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
        Try
            _useCombination = CBool(_activeJob.Trigger.TriggerInfoXml.SelectSingleNode("watcherbase/@combination").Value)
        Catch
            _useCombination = False
        End Try

        '--- de-serialise the info from the xml
        For Each ndWatcher As Xml.XmlNode In _activeJob.Trigger.TriggerInfoXml.SelectNodes("descendant::watcher")

            '--- create a SetupInfo for each watcher
            Dim aWatcherInfo As New WatcherInfo

            With ndWatcher
                aWatcherInfo.Name = .SelectSingleNode("@name").Value
                aWatcherInfo.QueueName = .SelectSingleNode("queue").InnerText
                aWatcherInfo.MessageSelector = .SelectSingleNode("messageselector").InnerText
                aWatcherInfo.MaxAgeUnit = .SelectSingleNode("maxtriggerage/@unit").InnerText
                aWatcherInfo.MaxAgeAmount = CInt(.SelectSingleNode("maxtriggerage/@amount").InnerText)
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
                .MessagesFoundUpToLastTime = New Hashtable

                '--- get a hashtable of messages on the queue
                Dim ht As Hashtable = MyMessenger.SnapshotQueue(.QueueName, .MessageSelector)


                '--- calculate the earliest Last Modified date we are interested in
                '--- (this is so we can throttle the number of triggers that we check)
                'All adjusted to be GMT to correspond with Tib timestamp
                Dim earliestMessageTimestamp As Date = GetAdjustedDate(DateTime.UtcNow, .MaxAgeUnit, .MaxAgeAmount * -1)

                Dim gunMessage As GunMessage = Nothing
                For Each de As DictionaryEntry In ht
                    gunMessage = CType(de.Key, GunMessage)

                    If CDate(gunMessage.CreatedAt) > earliestMessageTimestamp Then

                        '--- add to the cumulative files found hashtable if not there already,
                        '--- and add to triggers too
                        If Not (_messagesFoundUpToLastTime.ContainsKey(gunMessage.MessageID)) Then
                            .MessagesFoundUpToLastTime.Add(gunMessage.MessageID, gunMessage)

                            '--- create the Found Trigger info
                            Dim fti As New FoundTriggerInfo
                            fti.PulledTrigger = gunMessage
                            .TriggersFound.Add(gunMessage.MessageID, fti)

                        End If

                    End If

                Next

            End With

        Next 'aWatcherInfo


        '--- if anything found then create messages if all combination conditions met, etc
        If _combinationLogic = "and" Then

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

        ElseIf _combinationLogic = "or" Then

            '--- (assume not found)
            FoundValidTrigger = False

            '--- any watcher info objects can have trigger files
            For Each aWatcherInfo As WatcherInfo In _watcherInfoList

                '--- if there is a watcher with found files then flag is true and can exit
                If aWatcherInfo.TriggersFound.Count > 0 Then
                    FoundValidTrigger = True
                    Exit For
                End If

            Next

        End If


        '--- continue if at least 1 trigger found
        If FoundValidTrigger = True Then

            '--- populate the TriggersFound table
            _triggersFound = New Hashtable

            '--- for non-combination watchers carry on as usual
            If Not _useCombination Then

                For Each aWatcherInfo As WatcherInfo In _watcherInfoList

                    For Each de As DictionaryEntry In aWatcherInfo.TriggersFound

                        _triggersFound.Add(de.Key, de.Value)

                    Next

                Next aWatcherInfo

            ElseIf _useCombination Then

                '--- message objects
                Dim lhsMsg As GunMessage = Nothing
                Dim rhsMsg As GunMessage = Nothing

                '--- work out which of the watchers has the largest triggers found
                Dim lhs As WatcherInfo = CType(IIf(_watcherInfoList(0).TriggersFound.Count > _watcherInfoList(1).TriggersFound.Count, _watcherInfoList(0), _watcherInfoList(1)), WatcherInfo)
                Dim rhs As WatcherInfo = CType(IIf(lhs Is _watcherInfoList(0), _watcherInfoList(1), _watcherInfoList(0)), WatcherInfo)

                '--- loop through largest triggers found collection
                For Each lhsDE As DictionaryEntry In lhs.TriggersFound

                    '--- loop through the rhs GunMessages comapring the match fields
                    For Each rhsDE As DictionaryEntry In rhs.TriggersFound

                        '--- flag to hold match status
                        Dim fMatch As Boolean = False

                        '--- attempt to match each entry to the other list,
                        '--- based on the <match> nodes
                        For Each ndMatchDescription As XmlNode In _activeJob.JobXml.SelectNodes("descendant::matches/match")

                            '--- get the field names to match on
                            Dim lhsField As String = ndMatchDescription.SelectSingleNode("field[@watcher='" & lhs.Name & "']/@fieldname").InnerText
                            Dim rhsField As String = ndMatchDescription.SelectSingleNode("field[@watcher='" & rhs.Name & "']/@fieldname").InnerText

                            '--- get the gun messages from both sides
                            lhsMsg = CType(CType(lhsDE.Value, FoundTriggerInfo).PulledTrigger, GunMessage)
                            rhsMsg = CType(CType(rhsDE.Value, FoundTriggerInfo).PulledTrigger, GunMessage)

                            '--- test for match
                            fMatch = (lhsMsg.GetNamedValue(lhsField).ToString = rhsMsg.GetNamedValue(rhsField).ToString)

                            '--- can get out here if this one doesn't match
                            If Not fMatch Then Exit For

                        Next ndMatchDescription

                        '--- if a match then create the new message
                        If fMatch Then

                            Dim gmsg As GunMessage = MyMessenger.CreateDisconnectedMessage

                            '--- add new trigger info
                            For Each ndNewTriggerInfo As XmlNode In _activeJob.JobXml.SelectNodes("descendant::newtriggerinfo/pair")

                                Dim sTargetFieldName As String = ndNewTriggerInfo.SelectSingleNode("name").InnerText
                                Dim msgSource As GunMessage = CType(IIf(lhs.Name = ndNewTriggerInfo.SelectSingleNode("value/@sourcewatcher").Value, lhsMsg, rhsMsg), GunMessage)
                                Dim sSourceFieldName As String = ndNewTriggerInfo.SelectSingleNode("value/@field").Value
                                Dim sSourceValue As String = ""
                                Try
                                    sSourceValue = msgSource.GetNamedValue(sSourceFieldName).ToString
                                Catch ex As Exception
                                    sSourceValue = "## Field '" & sSourceFieldName & "' not found in source message ##"
                                End Try

                                gmsg.AddNameValuePair(sTargetFieldName, sSourceValue)

                            Next ndNewTriggerInfo

                            '--- add to the triggers found collection
                            Dim fti As New FoundTriggerInfo
                            fti.PulledTrigger = gmsg
                            _triggersFound.Add(CreateUniqueID, fti)

                        End If

                    Next rhsDE

                Next lhsDE

            End If


            'HIGH: Need to test for BC fields in the Trigger msg NVP, and then build the .Crumbler object on the FTI
            '--- need to check the messages for Crumbler info
            For Each de As DictionaryEntry In _triggersFound

                '--- extract the FTI and FI objects
                Dim fti As FoundTriggerInfo = CType(de.Value, FoundTriggerInfo)
                Dim gmsgTrigger As GunMessage = CType(fti.PulledTrigger, GunMessage)

                Try

                    Dim sTrailDescription As String, sTrailLabel As String, sTrailUid As String
                    Dim sNameNotFound As String = "#[NAME_NOT_FOUND]#"

                    '--- Trail Description
                    sTrailDescription = gmsgTrigger.GetNamedValue("BC_TrailDescription").ToString
                    
                    '--- Trail Label
                    sTrailLabel = gmsgTrigger.GetNamedValue("BC_TrailLabel").ToString

                    '--- Trail UID
                    sTrailUid = gmsgTrigger.GetNamedValue("BC_UID").ToString

                    '--- create the Crumbler if all values valid
                    fti.Crumbler = New Crumbler(gmsgTrigger.Pairs)

                Catch ex As Exception

                End Try

            Next


            '--- send messages to ATJ queue
            CreateMessages()

        End If

    End Sub



    ''' <summary>
    ''' helper sub-class to store settings for each watcher
    ''' </summary>
    ''' <remarks></remarks>
    Class WatcherInfo : Inherits WatcherBase.WatcherInfoBase
        Public QueueName As String
        Public MessageSelector As String
        Public MessagesFoundUpToLastTime As Hashtable
    End Class


End Class

