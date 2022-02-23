Imports OTIS
Imports System.IO
Imports System.Data.SqlClient
Imports System.Xml


Module RationaliserModule

#Region "Public Methods"

    Dim LatestTimestamp As Double
    Dim ConnectionString As String = ""
    Dim MyConnection As SqlClient.SqlConnection

    Public Sub Rationalise(ByVal MyMessenger As Messenger)

        ConnectionString = "Data Source=" & MyMessenger.Config.DBServer & ";Initial Catalog=" & MyMessenger.Config.DBName & ";Integrated Security=True"

        MyConnection = New SqlClient.SqlConnection(connectionString)

        ''Snapshot the whole queue
        'Dim ht As Hashtable = MyMessenger.SnapshotQueue(MyMessenger.Config.AllTriggeredJobsQ, , True)

        'Dim CountOfMessagesToRationalise As Integer = ht.Count
        Dim CountOfMovedMessages As Integer = 0
        Dim CountOfDestroyedMessages As Integer = 0
        Dim CountOfDatabaseMessages As Integer = 0

        ''Attempt to consume front message in snapshot from the queue
        'For Each de As DictionaryEntry In ht

        'Consume the front item from the AllJobs queue.  Quit if nothing found
        'Dim FrontMessage As GunMessage = MyMessenger.ConsumeMessage(MyMessenger.Config.AllTriggeredJobsQ, de.Value.ToString)

        Dim FrontMessage As GunMessage

        'If FrontMessage IsNot Nothing Then
        Do

            FrontMessage = MyMessenger.ConsumeMessage(MyMessenger.Config.AllTriggeredJobsQ)

            If FrontMessage IsNot Nothing Then

                'Make sure there are no more like it on this queue
                'RemoveDuplicatesOfThis(FrontMessage)

                Dim FrontMessageHash As String = FrontMessage.Hash

                'Check items in all downstream queues/DB for this hash value.
                If ExistsOnQueue(MyMessenger.Config.JobsInProgressQ, FrontMessageHash) Then
                    CountOfDestroyedMessages += 1
                ElseIf ExistsOnQueue(MyMessenger.Config.JobsDoneQ, FrontMessageHash) Then
                    CountOfDestroyedMessages += 1
                ElseIf ExistsOnQueue(MyMessenger.Config.QuarantineQ, FrontMessageHash) Then
                    CountOfDestroyedMessages += 1
                ElseIf ExistsInDatabase(FrontMessage.Hash) Then
                    CountOfDestroyedMessages += 1
                    CountOfDatabaseMessages += 1
                ElseIf Not ExistsOnQueue(MyMessenger.Config.JobsToDoQ, FrontMessageHash) Then

                    '--- *** CRUMBLER ***

                    '--- extract info for crumbler if available
                    MyCrumbler = New Crumbler(FrontMessage.Pairs)

                    '--- post the Crumb if all values valid
                    If MyCrumbler IsNot Nothing Then

                        '--- extract the info for this event
                        Dim xInfo As XmlNode = GetRationalisedInfo(FrontMessage)

                        MyCrumbler.DropCrumb("MovedToJobsToDo", xInfo)

                    End If


                    'If not found, post to the JobsToDo queue
                    MyMessenger.PostMessage(MyMessenger.Config.JobsToDoQ, FrontMessage)
                    CountOfMovedMessages += 1

                End If

                'End If

                'Next

            End If

        Loop Until FrontMessage Is Nothing

        'Dim ConsoleMessage As String = "Rationalised " & vbTab & CountOfMessagesToRationalise.ToString & vbCr
        'Console.WriteLine(ConsoleMessage)
        Dim ConsoleMessage As String = Format(Now, "HH:mm:ss") & vbCr
        Console.WriteLine(ConsoleMessage)
        ConsoleMessage = "Progressed " & vbTab & CountOfMovedMessages.ToString & vbCr
        Console.WriteLine(ConsoleMessage)
        ConsoleMessage = "Destroyed " & vbTab & (CountOfDestroyedMessages).ToString
        Console.WriteLine(ConsoleMessage)
        ConsoleMessage = "DB checked " & vbTab & (CountOfDatabaseMessages).ToString
        Console.WriteLine(ConsoleMessage)

        'Now make the latest snapshot of the JobsToDo queue unique.  
        UniquifyJobsToDo()

        'Check for orphaned JobsInProgress messages, move any back to AllTriggeredJobs
        CheckForOrphanedJobsInProgress()

        Console.WriteLine("")

    End Sub

    'Public Sub NewRationalise(ByVal MyMessenger As Messenger)

    '    'Snapshot AllTriggeredJobs messages that are newer than last logged timestamp
    '    Dim AllTriggeredJobs As String
    '    AllTriggeredJobs = MyMessenger.Config.AllTriggeredJobsQ
    '    Dim MsgColl As Microsoft.VisualBasic.Collection = MyMessenger.SnapshotQueue_Ordered(AllTriggeredJobs, "JMSTimestamp>" & LatestTimestamp.ToString)

    '    'The snapshot is time ordered.  Note time of newest snapped message because we will only ask for messages 
    '    'newer than this in subsequent snaps

    '    Do Until MsgColl.Count = 0

    '        LatestTimestamp = CType(MsgColl(MsgColl.Count), GunMessage).NumericTimestamp

    '        'Grab the first message on the queue 
    '        Dim CheckMessage As GunMessage = CType(MsgColl(1), GunMessage)
    '        Dim CheckHash As String = CheckMessage.Hash

    '        'Now check each of the other messages in the snapshot, attempt to remove dups
    '        For i As Integer = MsgColl.Count To 2 Step -1

    '            Dim gm As GunMessage = CType(MsgColl(i), GunMessage)

    '            If gm.Hash = CheckHash Then
    '                Try
    '                    MsgColl.Remove(i)
    '                    MyMessenger.ConsumeMessage(AllTriggeredJobs, gm.MessageID)
    '                Catch
    '                End Try
    '            End If

    '        Next

    '        'Check items in all downstream queues/DB for this hash value.
    '        If ExistsOnQueue(MyMessenger.Config.JobsInProgressQ, CheckHash) Then

    '        ElseIf ExistsOnQueue(MyMessenger.Config.JobsDoneQ, CheckHash) Then

    '        ElseIf ExistsOnQueue(MyMessenger.Config.QuarantineQ, CheckHash) Then

    '        ElseIf ExistsInDatabase(CheckHash) Then

    '        ElseIf Not ExistsOnQueue(MyMessenger.Config.JobsToDoQ, CheckHash) Then

    '            'If not found, post to the JobsToDo queue
    '            MyMessenger.PostMessage(MyMessenger.Config.JobsToDoQ, CheckMessage)

    '        End If

    '        '--- remove the message from collection so we don't act on it again, and we get to zero msgs eventually
    '        MsgColl.Remove(1)

    '    Loop

    '    'Now make the latest snapshot of the JobsToDo queue unique.  
    '    UniquifyJobsToDo()

    '    'Check for orphaned JobsInProgress messages, move any back to AllTriggeredJobs
    '    CheckForOrphanedJobsInProgress()

    'End Sub

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Browse a queue for a message with a specific hash
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="HashValue"></param>
    ''' <returns>true/false</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Private Function ExistsOnQueue(ByVal QueueName As String, _
                                    ByVal HashValue As String) As Boolean

        'Get the collection of gun messages on this queue that have our hash value
        Dim ht As Hashtable = MyMessenger.SnapshotQueue(QueueName, "hash=" & Qt(HashValue), True)

        'If the collection was empty, the item's not on the queue; otherwise it was
        If ht.Count = 0 Then
            Return False
        Else
            Return True
        End If

    End Function

    ''' <summary>
    ''' Browse a database for a message with a specific hash
    ''' </summary>
    ''' <param name="HashValue"></param>
    ''' <returns>true/false</returns>
    ''' <remarks>SDCP 20-Nov-2006</remarks>
    Private Function ExistsInDatabase(ByVal HashValue As String) As Boolean

        '        Return True
        Dim reader As SqlDataReader = Nothing

        Try

            If MyConnection Is Nothing Then
                MyConnection = New SqlClient.SqlConnection(ConnectionString)
            End If

            If MyConnection.State = ConnectionState.Closed Then
                Console.WriteLine("Connection opened" & vbCr)
                MyConnection.Open()
            End If

            '--- create database connection
            'Using MyConnection 'As New SqlClient.SqlConnection(connectionString)

            '--- create the command
            Dim checkCommand As SqlCommand = New SqlCommand("Select * from RPM_Hashes where hash=@hash and QueueName=@QueueName", MyConnection)

            '--- run it
            With checkCommand
                .Parameters.Add("@hash", SqlDbType.VarChar, 255)
                .Parameters("@hash").Value = HashValue

                .Parameters.Add("@QueueName", SqlDbType.VarChar, 255)
                .Parameters("@QueueName").Value = MySwitches("environment").ToString & "_" & "JobsDone"

                '--- return true if record count > 0
                reader = .ExecuteReader()
                Return reader.HasRows

            End With

            'End Using

        Catch ex As Exception

            Dim sErrorMessage As String = "#Unexpected error in Rationaliser/RationaliserModule.vb/ExistsInDatabase" & vbCr
            sErrorMessage += ex.Message & vbCr
            sErrorMessage += ex.StackTrace & vbCr
            sErrorMessage += ex.Source

            MyMessenger.Log(Messenger.LogType.ErrorType, sErrorMessage, HashValue)
            Return True

        Finally
            If reader IsNot Nothing Then
                reader.Close()
            End If

        End Try

    End Function

    ''' <summary>
    ''' Check all messages on JobsInProgress.  Any where the messageexpiry field value is smaller than
    ''' the time now, consume and move back onto AllTriggeredJobs
    ''' </summary>
    ''' <remarks>ML</remarks>
    Private Sub CheckForOrphanedJobsInProgress()

        Dim CountOfOrphanedJobsInProgress As Integer = 0

        With MyMessenger

            'Grab all the jobs.  Iterate through them checking timestamp+expiry > TibNowTime
            Dim AllJobsOnInProgQueue As Hashtable = .SnapshotQueue(.Config.JobsInProgressQ)
            For Each de As DictionaryEntry In AllJobsOnInProgQueue

                Dim gm As GunMessage = CType(de.Key, GunMessage)
                Dim iMessageTimestamp As Double = CDbl(gm.NumericTimestamp)
                Dim iExpiry As Long = CLng(gm.Pairs("timeout"))

                Dim iTibTimeNow As Double = .TIBTime

                If iExpiry + iMessageTimestamp < iTibTimeNow Then

                    'Consume the expired messages
                    Dim OrphanedJob As GunMessage = Nothing
                    Try
                        OrphanedJob = .ConsumeMessage(.Config.JobsInProgressQ, gm.MessageID)
                        CountOfOrphanedJobsInProgress += 1
                    Catch ex As Exception
                    End Try

                    If Not OrphanedJob Is Nothing Then

                        '--- create the crumb if possible
                        MyCrumbler = Nothing
                        MyCrumbler = New Crumbler(OrphanedJob.Pairs)
                        If MyCrumbler IsNot Nothing Then
                            MyCrumbler.DropCrumb("OrphanedJob", GetOrphanedjobInfo(OrphanedJob))
                        End If

                        '--- move message back to ATJ queue
                        .PostMessage(.Config.AllTriggeredJobsQ, OrphanedJob)

                    End If

                End If

            Next

        End With

        Console.WriteLine("Orphaned " & vbTab & CountOfOrphanedJobsInProgress.ToString)

    End Sub

    ''' <summary>
    ''' Runner only picks up jobs that are at least 10 seconds old, so Rationaliser should have passed over 
    ''' the queue a number of times in that window removing any duplicate messages.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UniquifyJobsToDo()

        Dim col As Collection = MyMessenger.SnapshotQueue_Ordered(MyMessenger.Config.JobsToDoQ)
        Dim CountOfNonUniqueJobs As Integer = 0
        For i As Integer = 1 To col.Count
            For j As Integer = i + 1 To col.Count
                If CType(col(i), GunMessage).Hash = CType(col(j), GunMessage).Hash Then
                    'two jobs with same hash, so consume the newer one
                    Try
                        MyMessenger.ConsumeMessage(MyMessenger.Config.JobsToDoQ, CType(col(j), GunMessage).MessageID)
                        CountOfNonUniqueJobs += 1
                    Catch
                    End Try
                End If
            Next
        Next

        Console.WriteLine("Non-unique " & vbTab & CountOfNonUniqueJobs.ToString)

    End Sub

    ''' <summary>
    ''' Take a snapshot of the rest of AllJobs queue, compare to the front message.  Remove any dups
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RemoveDuplicatesOfThis(ByVal ComparisonMessage As GunMessage)

        'Get a collection of all messages on the queue with a matching hash
        Dim MsgSelector As String = "hash=" & Qt(ComparisonMessage.Hash)
        Dim ht As Hashtable = MyMessenger.SnapshotQueue(MyMessenger.Config.AllTriggeredJobsQ, MsgSelector)

        'Consume them all
        For Each de As DictionaryEntry In ht

            Try
                MyMessenger.ConsumeMessage(MyMessenger.Config.AllTriggeredJobsQ, CType(de.Key, GunMessage).MessageID)
            Catch
            End Try

            ''--- get the message
            'Dim TestGunMessage As GunMessage = CType(de.Key, GunMessage)

            ''Test its' hash against ComparisonMessage hash
            'If TestGunMessage.Hash.ToLower = ComparisonMessage.Hash.ToLower Then

            '    'Consume this message if it matches
            '    MyMessenger.ConsumeMessage(MyMessenger.Config.AllTriggeredJobsQ, TestGunMessage.MessageID)

            'End If

        Next

    End Sub


    ''' <summary>
    ''' Gets info for rationalised job to give to the Crumb
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <returns></returns>
    ''' <remarks>SDCP 31-Jan-2008</remarks>
    Private Function GetRationalisedInfo(ByVal Message As GunMessage) As XmlNode
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


        '--- return the node
        Return xDoc.DocumentElement

    End Function


    ''' <summary>
    ''' Gets info for orphaned job to give to the Crumb
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <returns></returns>
    ''' <remarks>SDCP 31-Jan-2008</remarks>
    Private Function GetOrphanedjobInfo(ByVal Message As GunMessage) As XmlNode

        Return GetRationalisedInfo(Message)

    End Function


#End Region

End Module
