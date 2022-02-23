Module MainModule

    Sub Main()

        Console.WriteLine("This is Rationaliser")

        'Write some stuff to the console window
        Dim AppPath As String = Environment.GetCommandLineArgs(0)
        Console.Write(AppPath & vbLf)
        Console.Write("Built: " & Format(FileSystem.FileDateTime(AppPath), "ddMMMyyyy HH:mm") & vbLf)

        '--- get command-line switches
        MySwitches = New Switches

        Dim sMessage As String = ""
        Dim sConfigFile As String = GetConfigFileName()

        MyMessenger = SetUpForMessaging(sConfigFile)
        If MyMessenger Is Nothing Then
            Exit Sub
        End If

        'Accessor for the config obj
        Dim ConfigObject As Config = MyMessenger.Config

        'Prepare a status message object
        Dim myStatusMessage As StatusMessage = New StatusMessage()

        'Keep looping.  If a hari kari message is found, application will terminate
        Do
            Try
                'Do any status related stuff
                myStatusMessage.Check()

                'Rationalise
                Rationalise(MyMessenger)

                'Go to sleep for some secs
                Dim iSleep As Integer = MyMessenger.Config.PollFrequency
                Threading.Thread.Sleep(iSleep * 1000)

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
