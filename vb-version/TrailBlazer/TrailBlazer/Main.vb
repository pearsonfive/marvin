Imports OTIS
Imports System.Data.SqlClient


Module Main

    Public Sub Main()

        Dim lInsertedID As Long

        '--- get the command line switches
        MySwitches = New Switches

        Dim sMessage As String = ""
        Dim sConfigFile As String = GetConfigFileName()

        'If unable to setup messaging, log to file & quit
        MyMessenger = SetUpForMessaging(GetConfigFileName)
        If MyMessenger Is Nothing Then
            End
        End If

        'Write some stuff to the console window
        Dim AppPath As String = Environment.GetCommandLineArgs(0)
        Console.Write(AppPath & vbLf)
        Console.Write("Built: " & Format(FileSystem.FileDateTime(AppPath), "ddMMMyyyy HH:mm") & vbLf)
        Console.Write("Environment: " & MySwitches("environment") & vbLf)

        Do

            Try

                '--- look for messages on the Trail queue
                Dim msgCollection As Hashtable = MyMessenger.SnapshotQueue(MyMessenger.Config.BreadcrumbsQ)

                '--- if we have a message then process it
                If msgCollection.Count > 0 Then

                    '--- create database connection
                    '            Dim connectionString As String = "Data Source=" & MyMessenger.Config.DBServer & ";Initial Catalog=" & MyMessenger.Config.DBName & ";Integrated Security=True"
                    Dim connectionString As String = "Data Source=" & MyMessenger.Config.DBServer & ";Initial Catalog=" & MyMessenger.Config.DBName & ";Integrated Security=True"
                    Using connection As New SqlClient.SqlConnection(connectionString)

                        '--- open the connection
                        connection.Open()

                        Dim adapter As New SqlDataAdapter()

                        '--- create the insert commands
                        Dim insertMapCommand As SqlCommand = New SqlCommand _
                        ("INSERT INTO MarvinTrailMap(Environment, TrailLabel, TrailDescription, UID, EventName, JMSTimestamp, CalculatedTimestamp) values(@Environment, @TrailLabel, @TrailDescription, @UID, @EventName, @JMSTimestamp, @CalculatedTimestamp); select cast(scope_identity() as int);", connection)

                        Dim insertInfoCommand As SqlCommand = New SqlCommand _
                        ("INSERT INTO MarvinTrailInfo(Info, TrailMapID) values(@Info, @TrailMapID)", connection)

                        '--- add the parameters for the insert commands
                        insertMapCommand.Parameters.Add("@Environment", SqlDbType.VarChar, 255, "environment")
                        insertMapCommand.Parameters.Add("@TrailLabel", SqlDbType.VarChar, 255, "traillabel")
                        insertMapCommand.Parameters.Add("@TrailDescription", SqlDbType.VarChar, 255, "traildescription")
                        insertMapCommand.Parameters.Add("@UID", SqlDbType.VarChar, 50, "uid")
                        insertMapCommand.Parameters.Add("@EventName", SqlDbType.VarChar, 255, "eventname")
                        insertMapCommand.Parameters.Add("@JMSTimestamp", SqlDbType.BigInt, 4, "jmstimestamp")
                        insertMapCommand.Parameters.Add("@CalculatedTimestamp", SqlDbType.DateTime, 8, "calculatedtimestamp")

                        insertInfoCommand.Parameters.Add("@Info", SqlDbType.Text, 2147483647, "info")
                        insertInfoCommand.Parameters.Add("@TrailMapID", SqlDbType.Int, 4, "trailmapid")

                        '--- loop through messages in the collection and post to DB
                        For Each de As DictionaryEntry In msgCollection

                            '--- get the message
                            Dim gunMessage As GunMessage = CType(de.Key, GunMessage)

                            '--- extract the MarvinTrailMap table stuff
                            Dim Environment As String = MySwitches("environment")
                            Dim TrailLabel As String = gunMessage.GetNamedValue("BC_TrailLabel").ToString
                            Dim TrailDescription As String = gunMessage.GetNamedValue("BC_TrailDescription").ToString
                            Dim UID As String = gunMessage.GetNamedValue("BC_UID").ToString
                            Dim EventName As String = gunMessage.GetNamedValue("BC_EventName").ToString
                            Dim JMSTimestamp As Long = gunMessage.NumericTimestamp
                            Dim CalculatedTimestamp As Date = TIBTimeToWindowsTime(JMSTimestamp)

                            '--- create entry in the MarvinTrailMap table
                            With insertMapCommand
                                .Parameters("@Environment").Value = Environment
                                .Parameters("@TrailLabel").Value = TrailLabel
                                .Parameters("@TrailDescription").Value = TrailDescription
                                .Parameters("@UID").Value = UID
                                .Parameters("@EventName").Value = EventName
                                .Parameters("@JMSTimestamp").Value = JMSTimestamp
                                .Parameters("@CalculatedTimestamp").Value = CalculatedTimestamp
                                Try
                                    lInsertedID = Convert.ToInt32(.ExecuteScalar())
                                Catch e As Exception
                                    Continue For
                                End Try
                            End With


                            If lInsertedID <> 0 Then

                                '--- create payload stuff
                                Dim Info As String = gunMessage.GetNamedValue("BC_Info").ToString

                                '--- create entry in MarvinTrailInfo table
                                With insertInfoCommand
                                    .Parameters("@Info").Value = Info
                                    .Parameters("@TrailMapID").Value = lInsertedID
                                    Try
                                        .ExecuteNonQuery()
                                    Catch e As Exception
                                        Continue For
                                    End Try
                                End With


                                '--- consume the message off the queue
                                MyMessenger.ConsumeMessage(gunMessage.QueueName, gunMessage.MessageID)

                            Else

                            End If

                        Next

                    End Using

                End If

                '--- sleep for n seconds
                Threading.Thread.Sleep(30 * 1000)

            Catch ex As Exception
                If MyMessenger Is Nothing Then
                    'If it errored because it didn't have a config, re-initialise
                    MyMessenger = New Messenger(sConfigFile, sTHIS_APP)
                Else
                    sMessage = "#Unexpected error in " & ex.Source & vbCr
                    sMessage += ex.StackTrace
                    MyMessenger.Log(Messenger.LogType.ErrorType, sMessage)
                    Exit Sub
                End If
            End Try

        Loop

    End Sub


End Module
