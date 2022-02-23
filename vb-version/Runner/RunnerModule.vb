Imports OTIS
Imports System.IO
Imports System.Xml


Module RunnerModule

    ''' <summary>
    ''' Reporting status at all points, set up the atom, then run it.  If the atom timed out, re-post
    ''' the job to JobsToDo.  Otherwise, check that the atom posted a message to JobsDone.  If it
    ''' didn't, re-post to JobsToDo.  Otherwise, job finished successfully.
    ''' </summary>
    ''' <param name="gunMessage"></param>
    ''' <param name="inProgressMessageId"></param>
    ''' <remarks></remarks>
    Sub DoJob(ByVal gunMessage As GunMessage, _
                    ByVal inProgressMessageId As String, _
                    ByVal myAtoms As Atoms)

        'pick up the payload, hash and jobname.  If any of these are missing, quit
        Dim sPayload As String = gunMessage.Payload
        Dim sHash As String = gunMessage.Hash
        Dim sJobname As String = gunMessage.JobName
        Dim sMessage As String = ""

        If sPayload = "" Or sHash = "" Or sJobname = "" Then
            sMessage = "#Error: Invalid message found" & vbCr
            sMessage = sMessage & "hash=" & sHash & ";"
            sMessage = sMessage & "jobname=" & sJobname & ";"
            sMessage = sMessage & "payload=" & sPayload
            MyMessenger.Log(Messenger.LogType.ErrorType, sMessage)
            Exit Sub
        End If

        'Now write default (from main config) & override (from job payload) settings
        If myAtoms.Errors.Count > 0 Then

            'Unable to load all settings.  Write message and quit
            sMessage = "#Unable to load all settings" & vbCr
            For Each DE As DictionaryEntry In CType(DE.Value, Hashtable)
                sMessage += DE.Value & vbCr
            Next

            MyMessenger.Log(Messenger.LogType.ErrorType, sMessage, sHash)
            Exit Sub

        End If

        ''Report status at start of job
        'sMessage = "Starting " & sJobname
        'MyMessenger.Log(Messenger.LogType.StatusType, sMessage, sHash)

        'Now run the atom, if hash is still unique on the JobsInProgress queue
        '--- browse ahead on JobsInProgress queue to see if another runner has picked it up already
        If MyMessenger.SnapshotQueue(MyMessenger.Config.JobsInProgressQ, "hash=" & Qt(gunMessage.Hash)).Count = 1 Then

            'Grab a start time for the atom
            Dim clock As GunMessage = MyMessenger.BroadcastMessage(MyMessenger.Config.ClockTopic, "")
            Dim StartTime As String = "JMSTimestamp>" & gunMessage.NumericTimestamp

            Dim fHasRunProperly As Boolean = myAtoms.Run()

            If Not fHasRunProperly Then
                'Report status at if job didn't complete successfully
                sMessage = "<statusmessage>Finished " & sJobname & IIf(Not fHasRunProperly, " with errors", "") & "</statusmessage>"
                'sMessage += MyMessenger.AtomInfo
                MyMessenger.Log(Messenger.LogType.StatusType, sMessage, sHash)
            End If


            'Runner relies on the application itself to post to the "done" queue.  Check this was done - if not, 
            're-pub to AllTriggeredJobs
            Dim gmsg As GunMessage = MyMessenger.ConsumeMessage(MyMessenger.Config.JobsInProgressQ, inProgressMessageId)

            '--- if the hash does not exist on JobsInProgress any more for some reason, then End here
            '--- as Rationaliser will move the message to AllTriggeredJobs in a few seconds any way
            If gmsg Is Nothing Then

                sMessage = "<statusmessage>JobsInProgress message has disappeared</statusmessage>"
                sMessage += MyMessenger.AtomInfo
                MyMessenger.Log(Messenger.LogType.StatusType, sMessage, sHash)
                End

            End If
            
            '--- check that Done msg count = Atoms count for this hash
            '--- otherwise post back to AllTriggeredJobs
            Dim sMessageSelector As String = "hash=" & Qt(sHash)
            If MyMessenger.SnapshotQueue(MyMessenger.Config.JobsDoneQ, sMessageSelector).Count _
                    <> myAtoms.Atoms.Count Then

                'Atom didn't run successfully.  Either put back to start of process, or quarantine.  Either
                'way, pick up any new status messages relating to this job and add to audit

                'Write a messages node to be appended to audit
                Dim NewMessages As String = ""
                sMessageSelector += " AND " & StartTime
                Dim ht As Hashtable = MyMessenger.SnapshotQueue(MyMessenger.Config.StatusQ, sMessageSelector)
                For Each de As DictionaryEntry In ht
                    Dim StatusMessage As GunMessage = CType(de.Key, GunMessage)
                    NewMessages += "<message>" & StatusMessage.GetNamedValue("message") & "</message>"
                Next

                If CType(gmsg.GetNamedValue("auditcount"), Integer) > MyMessenger.Config.MaxAuditPoints Then
                    'pub gmsg to quarantine queue
                    Call MyMessenger.PostMessage(MyMessenger.Config.QuarantineQ, "", "", "", gmsg.Pairs, , NewMessages)
                Else
                    'pub gmsg back to AllTriggeredJobs queue
                    Call MyMessenger.PostMessage(MyMessenger.Config.AllTriggeredJobsQ, "", "", "", gmsg.Pairs, , NewMessages)
                End If

            End If

        Else
            '--- There is a problem if >1 matching hashes are on JobsInProgress so,
            '--- consume my message and post back to AllTriggeredJobs.  
            '--- This means it will be rationalised again if required.

            '-- consume first
            Dim oldInProgressMessage As GunMessage = MyMessenger.ConsumeMessage(MyMessenger.Config.JobsInProgressQ, inProgressMessageId)

            '--- publish back to AllTriggeredJobs
            Try
                Call MyMessenger.PostMessage(MyMessenger.Config.AllTriggeredJobsQ, oldInProgressMessage)
            Catch ex As Exception
            End Try

        End If

    End Sub



    ''' <summary>
    ''' Gets info for rationalised job to give to the Crumb
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <returns></returns>
    ''' <remarks>SDCP 31-Jan-2008</remarks>
    Public Function GetBaseRunnerInfo(ByVal Message As GunMessage) As XmlNode
        Dim xDoc As XmlDocument
        Dim xNode As XmlNode

        '--- create the doc
        xDoc = New XmlDocument
        xDoc.LoadXml("<info/>")

        '--- start adding nodes:

        '--- machine name
        xNode = xDoc.CreateElement("machine")
        xNode.InnerXml = MachineName()
        xDoc.DocumentElement.AppendChild(xNode)

        '--- pid
        xNode = xDoc.CreateElement("pid")
        xNode.InnerXml = CurrentPID().ToString
        xDoc.DocumentElement.AppendChild(xNode)

        '--- app name
        xNode = xDoc.CreateElement("app")
        xNode.InnerXml = My.Application.Info.AssemblyName
        xDoc.DocumentElement.AppendChild(xNode)

        '--- path
        xNode = xDoc.CreateElement("path")
        xNode.InnerXml = My.Application.Info.DirectoryPath
        xDoc.DocumentElement.AppendChild(xNode)

        '--- original jms msg id
        xNode = xDoc.CreateElement("originaljmsmessageid")
        xNode.InnerXml = Message.MessageID
        xDoc.DocumentElement.AppendChild(xNode)

        '--- return the data
        Return xDoc.DocumentElement

    End Function


End Module
