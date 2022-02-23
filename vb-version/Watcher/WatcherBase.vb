Public MustInherit Class WatcherBase



    Protected _triggersFound As Hashtable
    Protected _jobDescriptionXml As Xml.XmlDocument
    Protected _activeJob As Job
    Protected _foundValidTrigger As Boolean
    Protected _lastPollDateTime As DateTime
    '    Protected _pollInterval As Integer
    Protected _pollFrequencyUnit As String
    Protected _pollFrequencyAmount As Integer
    Protected _savePollDateTime As DateTime
    Protected _stopWatcher As Boolean

    Protected _holidays As String = ""



    ''' <summary>
    ''' custom event to raise when just about to check for trigger
    ''' </summary>
    ''' <param name="Watcher"></param>
    ''' <remarks>SDCP 26-Sep-2006</remarks>
    Public Event BeforeCheckingForTrigger(ByVal Watcher As WatcherBase)

    ''' <summary>
    ''' custom event to raise when just checked for trigger
    ''' </summary>
    ''' <param name="Watcher"></param>
    ''' <remarks>SDCP 26-Sep-2006</remarks>
    Public Event AfterCheckingForTrigger(ByVal Watcher As WatcherBase)

    ''' <summary>
    ''' Actually does the polling
    ''' </summary>
    ''' <remarks>SDCP 06-Oct-2006</remarks>
    Protected MustOverride Sub Poller()

    ''' <summary>
    ''' extract setup info from the job definition
    ''' </summary>
    ''' <remarks>SDCP 01-Nov-2006</remarks>
    Protected MustOverride Sub ExtractSetupInfo()


    ''' <summary>
    ''' the current trigger object
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ActiveJob() As Job
        Get
            Return _activeJob
        End Get
        Set(ByVal value As Job)
            _activeJob = value
        End Set
    End Property


    ''' <summary>
    ''' the expanded (de-tokenised) job description XML from the job description file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property JobDescriptionXml() As Xml.XmlDocument
        Get
            Return _jobDescriptionXml
        End Get
        Set(ByVal value As Xml.XmlDocument)
            _jobDescriptionXml = value
        End Set
    End Property


    ''' <summary>
    ''' collection of possible triggers found
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TriggersFound() As Hashtable
        Get
            Return _triggersFound
        End Get
    End Property


    ''' <summary>
    ''' when a valid trigger has been found
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FoundValidTrigger() As Boolean
        Get
            Return _foundValidTrigger
        End Get
        Set(ByVal value As Boolean)
            _foundValidTrigger = value
        End Set
    End Property


    ''' <summary>
    ''' Records the last time the watcher was polled
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 03-Oct-2006</remarks>
    Public Property LastPollDateTime() As DateTime
        Get
            Return _lastPollDateTime
        End Get
        Set(ByVal value As DateTime)
            _lastPollDateTime = value
        End Set
    End Property


    ''' <summary>
    ''' method that actually runs the check for the trigger
    ''' </summary>
    ''' <remarks>SDCP 26-Sep-2006</remarks>
    Protected Sub StartPolling()

        '--- if the PollFrequencyUnit is "once" then this is a special case where we poll just once
        '--- (e.g. for TopicWatcher we poll once, and set up a listener, so no need to poll again)
        '--- (can't really think of another use case for PollFrequencyUnit = "once")

        If _pollFrequencyUnit = "once" Then

            '--- just poll
            Poller()

        Else

            '--- create poll interval
            'Dim period As New TimeSpan(0, 0, _pollInterval)
            Dim myNow As DateTime = Now
            Dim period As TimeSpan = GetAdjustedDate(myNow, _pollFrequencyUnit, _pollFrequencyAmount) - myNow

            Do

                If InTemporalScope(_activeJob) Then

                    '--- remember the current time
                    _savePollDateTime = Now

                    '--- poll for triggers
                    Poller()

                    '--- we now needto set the last poll time
                    _lastPollDateTime = _savePollDateTime

                End If

                '--- sleep until next poll time
                Threading.Thread.Sleep(period)

            Loop Until _stopWatcher = True

        End If

    End Sub




    ''' <summary>
    ''' Creates the messages on the AllTriggeredJobs queue
    ''' </summary>
    ''' <remarks>SDCP 01-Nov-2006</remarks>
    Protected Sub CreateMessages()
        Dim sHash As String
        Dim xmlDoc As New Xml.XmlDocument
        Dim sJobGroup As String = ""

        For Each FoundTrigger As FoundTriggerInfo In _triggersFound.Values

            '--- set the trigger timestamp
            Me.ActiveJob.Trigger.Timestamp = Now
            Me.ActiveJob.Trigger.UniqueID = CreateUniqueID()

            '--- create the payload
            Dim payloadCreator As New PayloadCreator(Me, FoundTrigger.PulledTrigger)
            xmlDoc.LoadXml(payloadCreator.GetPayload())
            Me._activeJob.Trigger.ExpandedTriggerInfoXml = xmlDoc.DocumentElement

            '--- create the hash 
            Dim hashCreator As New HashCreator(Me)
            sHash = hashCreator.GetHash()


            '--- expand again just for #[HASH]#
            Dim sCurrentPayload As String = Me._activeJob.Trigger.ExpandedTriggerInfoXml.OuterXml
            sCurrentPayload = sCurrentPayload.Replace("#[HASH]#", sHash)

            xmlDoc.LoadXml(sCurrentPayload)
            Me._activeJob.Trigger.ExpandedTriggerInfoXml = xmlDoc.DocumentElement


            'xHIGH: *** BrowseAhead is off in WatcherBase ***
            '--- browse ahead on queue to see if this hash exists yet
            If Not MyMessenger.HashExistsOnQueue(sHash, MyMessenger.Config.AllTriggeredJobsQ) Then

                Dim htNameValuePair As New Hashtable

                htNameValuePair.Add("payload", Me._activeJob.Trigger.ExpandedTriggerInfoXml.OuterXml)
                htNameValuePair.Add("hash", sHash)


                '--- CRUMBLER
                '--- we will need to create one if we are not the top of the chain
                If FoundTrigger.Crumbler Is Nothing Then

                    '--- create crumbler object
                    FoundTrigger.Crumbler = New Crumbler(Me._activeJob.Trigger.ExpandedTriggerInfoXml)

                End If

                '--- add the NVPs for the crumbler stuff
                htNameValuePair.Add("BC_TrailLabel", FoundTrigger.Crumbler.TrailLabel)
                htNameValuePair.Add("BC_TrailDescription", FoundTrigger.Crumbler.TrailDescription)
                htNameValuePair.Add("BC_UID", FoundTrigger.Crumbler.UID)

                '--- send the message:

                '--- try and work out the job group
                sJobGroup = ""
                Try
                    sJobGroup = Me._activeJob.JobXml.ParentNode.Attributes("name").Value
                Catch ex As Exception
                    sJobGroup = ""
                End Try
                MyMessenger.PostMessage(MyMessenger.Config.AllTriggeredJobsQ, sHash, Me._activeJob.Trigger.ExpandedTriggerInfoXml.OuterXml, Me._activeJob.Name, htNameValuePair, sJobGroup)

                'TESTING: test output of trail messages
                'FoundTrigger.Crumbler.DropCrumb("MessageWrittenToAllTriggeredJobs_ForTestingPurposesOnly", Me._activeJob.Trigger.ExpandedTriggerInfoXml.SelectSingleNode("trigger/hashbuilder"))

            End If

        Next FoundTrigger

    End Sub


    ''' <summary>
    ''' check that the trigger is within its run window, i.e. not a Sunday etc
    ''' </summary>
    ''' <param name="CurrentJob">The job being tested</param>
    ''' <returns>True if in scope, False if not</returns>
    ''' <remarks>SDCP 26-Sep-2006</remarks>
    Private Function InTemporalScope(ByVal CurrentJob As Job) As Boolean

        '--- get the current time and day
        Dim currentTime As Date = Now
        Dim currentDay As String = [Enum].GetName(GetType(System.DayOfWeek), Now.DayOfWeek)

        '--- check to see if a holiday
        If InStr(1, _holidays, currentTime.ToString("dd-MMM-yyyy"), CompareMethod.Text) <> 0 Then
            Return False
        End If

        '--- get the timings node
        Dim xTimings As Xml.XmlNode = CurrentJob.JobXml.SelectSingleNode("descendant::timings")

        '--- loop through each time block to check
        For Each xTimeBlock As Xml.XmlNode In xTimings.SelectNodes("descendant::timeblock")

            '--- extract the start and end times
            Dim startTime As Date, endTime As Date
            Date.TryParse(xTimeBlock.Attributes("start").Value, startTime)
            Date.TryParse(xTimeBlock.Attributes("end").Value, endTime)

            '--- test that we are between the start and end times
            If (Now > startTime And Now < endTime) Then

                '--- test we are in one of the valid days
                For Each xDay As Xml.XmlNode In xTimeBlock.SelectNodes("descendant::day")
                    If xDay.InnerText = currentDay Then Return True
                Next

            End If

        Next

        Return False

    End Function


    Protected Sub LoadHolidays()
        Dim sHolidayIdentifier As String
        Dim fSkipHolidays As Boolean
        Dim sHolidaysFilename As String = ""

        '--- see if we are interested in holidays
        Try
            fSkipHolidays = (_activeJob.JobXml.SelectSingleNode("descendant::timings").Attributes("skipholidays").Value = "true")
        Catch ex As Exception
            fSkipHolidays = False
        End Try

        '--- only need to worry if skip holidays is on
        If fSkipHolidays Then

            '--- get the holiday identifier
            sHolidayIdentifier = _activeJob.JobXml.SelectSingleNode("descendant::timings").Attributes("holidays").Value

            '--- get the filename
            For Each de As DictionaryEntry In MyMessenger.Config.HolidayFiles
                If de.Key.ToString = sHolidayIdentifier Then
                    sHolidaysFilename = de.Value.ToString()
                    Exit For
                End If

                '--- identifier not in config file, so exit
                Exit Sub

            Next

            '--- load the holidays file
            Try
                _holidays = IO.File.OpenText(sHolidaysFilename).ReadToEnd
            Catch ex As Exception
            End Try

        End If

    End Sub



    ''' <summary>
    ''' helper sub-class to store settings for each watcher
    ''' </summary>
    ''' <remarks></remarks>
    Class WatcherInfoBase

        Public Name As String
        Public MaxAgeUnit As String
        Public MaxAgeAmount As Integer
        Public TriggersFound As Hashtable

    End Class


    Class FoundTriggerInfo

        '--- a Crumbler object attached to the trigger
        Public Crumbler As Crumbler

        '--- the actual trigger object, e.g. a FileInfo, or GunMessage
        Public PulledTrigger As Object

    End Class


End Class
