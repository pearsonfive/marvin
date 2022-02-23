Imports OTIS

Module Main

    ''' <summary>
    ''' Standard initialisation stuff.  Then run a check for status messages; pick a message off the
    ''' JobsToDo queue that's at least 10 secs old; if a message with the same hash isn't on JobsInProgress 
    ''' post it there; spawn a new Runner instance; run the job
    ''' </summary>
    ''' <remarks></remarks>
    Sub Main()

        Console.WriteLine("This is Runner")

        '--- get command-line switches
        MySwitches = New Switches

        'determine the location of this exe and pick up the config from switch
        Dim sRunnerExeLocation As String = Environment.GetCommandLineArgs(0)

        'Set up all the TIBCO stuff
        MyMessenger = SetUpForMessaging(GetConfigFileName)
        If MyMessenger Is Nothing Then
            End
        End If

        'Setup statusmessages object
        Dim myStatusMessage As StatusMessage = New StatusMessage()

        'Keep spinning until finds a message to act upon
        Do

            Dim sMsgSelector As String = "", gunMessage As GunMessage = Nothing

            'Big try/catch, should only be for truly catastrophic
            Try

                'Check and action any admin tasks 
                myStatusMessage.Check()

                'Don't proceed if the box is already heavily loaded
                Do Until CPUUsed() < 90
                    System.Threading.Thread.Sleep(MyMessenger.Config.PollFrequency * 1000)
                Loop

                'Get tibco time by looking at a broadcast message's timestamp
                'Attempt to consume a message at least 10 seconds old from the queue. 
                Dim iTibTimeNow As Double = MyMessenger.TIBTime
                sMsgSelector = "JMSTimestamp<" & iTibTimeNow - (MyMessenger.Config.RunnerMessageAge * 1000)
                gunMessage = MyMessenger.ConsumeMessageBySelector(MyMessenger.Config.JobsToDoQ, sMsgSelector)

                '--- if we have a message then process it
                If gunMessage IsNot Nothing Then

                    With MyMessenger

                        '--- browse ahead on JobsInProgress queue to see if another runner has picked it up already
                        sMsgSelector = "hash=" & Qt(gunMessage.Hash)
                        If .SnapshotQueue(.Config.JobsInProgressQ, sMsgSelector).Count = 0 Then


                            'Setup the atom, given the gun message that triggered the job. 
                            Dim myAtoms As Atoms = New Atoms(gunMessage, .Config.DomDoc)

                            If myAtoms.Warnings.ContainsKey("Null") Then

                                'Special case - recognise "null" as "no atom to run".  Post message to "done", pause, then look for
                                'another job
                                .PostMessage(.Config.JobsDoneQ, gunMessage)

                                '--- sleep for some time
                                System.Threading.Thread.Sleep(MyMessenger.Config.PollFrequency * 1000)

                            Else

                                'Add a "timeout" field to the message - timeout is in millisecs.  Add 10 sec because message is
                                'at least 10 secs old at this point
                                Dim iTimeout As Long = (10 + myAtoms.Timeout) * 1000

                                Dim nvp As Hashtable = gunMessage.Pairs
                                Try
                                    nvp.Add("timeout", iTimeout)
                                Catch
                                    nvp.Remove("timeout")
                                    nvp.Add("timeout", iTimeout)

                                End Try

                                '--- put the message on the InProgress queue so that it exists somewhere in the framework
                                Dim inProgressMessageId As String
                                inProgressMessageId = .PostMessage(.Config.JobsInProgressQ, "", "", "", nvp)

                                '--- at this point we have a message that only this Runner is going to be handling
                                '--- so can start Crumbling:

                                '--- get info from NVP
                                MyCrumbler = New Crumbler(nvp)

                                '--- drop a crumb to show we are starting to run this job
                                MyCrumbler.DropCrumb("TakeOwnershipOfJob", GetBaseRunnerInfo(gunMessage))

                                'Don't spawn a new instance while de-bugging
                                If InStr(1, sRunnerExeLocation, "vshost.exe", CompareMethod.Text) = 0 Then

                                    'Spawn a new instance of runner that will do the next job
                                    Shell(Environment.CommandLine, AppWinStyle.Hide, False)
                                    System.Threading.Thread.Sleep(MyMessenger.Config.PollFrequency * 1000)
                                End If

                                'run a job from the queue
                                DoJob(gunMessage, inProgressMessageId, myAtoms)

                                '--- and now die off
                                End

                            End If

                        End If

                    End With

                End If

                '--- wait for a bit
                System.Threading.Thread.Sleep(MyMessenger.Config.PollFrequency * 1000)

            Catch ex As Exception

                'Set up for chatting
                MyChatter = New Chatter
                MyChatter.ReadChatSettings()

                Dim sMessage As String = MyChatter.RunLabel & " ::ERROR:: " & vbCr & vbCr
                sMessage += "Environment: " & MySwitches("environment") & vbCr & vbCr
                sMessage += "MsgSelector: " & sMsgSelector & vbCr & vbCr
                sMessage += "Machine: " & MachineName() & vbCr & vbCr
                sMessage += ex.Message & vbCr & vbCr
                sMessage += ex.StackTrace & vbCr & vbCr
                sMessage += "Closing this Runner and spawning a new instance"
                MyChatter.Chat(sMessage)

                LastDitchLog(sMessage)

                End

            End Try

        Loop

    End Sub

End Module
