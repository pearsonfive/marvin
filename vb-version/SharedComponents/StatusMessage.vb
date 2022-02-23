Imports OTIS

Public Class StatusMessage

#Region "Private"

    Private _nvp As Hashtable
    Private _MessageID As String
    '    Private _StatusType As MessageType
    '    Private _StatusQueue As String
    Private _MessageSelector As String

    Private Sub CheckForHariKiri()

        'Construct msg selector
        Dim MsgSelector As String = Replace(_MessageSelector, "#[StatusType]#", "harakiri")

        'If any messages like this are returned, app should die
        Dim ht As Hashtable
        ht = MyMessenger.SnapshotQueue(MyMessenger.Config.StatusQ, MsgSelector)
        If ht.Count > 0 Then
            End
        End If

    End Sub

    Private Sub CheckForWhoIsOutThere()

        'Construct msg selector
        Dim MsgSelector As String = Replace(_MessageSelector, "#[StatusType]#", "whoisoutthere")

        'If any messages returned send an "I Am Here" message
        Dim ht As Hashtable
        ht = MyMessenger.SnapshotQueue(MyMessenger.Config.StatusQ, MsgSelector)
        If ht.Count > 0 Then
            Dim nvp As Hashtable = GetBasicFields()
            nvp.Add("iamhere", Format(Now, "yyyyMMdd hh:mm:ss"))

            MyMessenger.PostMessage(MyMessenger.Config.StatusQ, "", "", "", nvp)
        End If

    End Sub

    Private Sub CheckForAppConfigReload()

        'Construct msg selector
        Dim MsgSelector As String = Replace(_MessageSelector, "#[StatusType]#", "appconfigreload")

        'If any messages returned, re-set the app config
        Dim ht As Hashtable
        ht = MyMessenger.SnapshotQueue(MyMessenger.Config.StatusQ, MsgSelector)

        '--- run it again and kill it off
        If ht.Count > 0 Then

            '--- consume the message so it doesn't trigger this again
            For Each de As DictionaryEntry In ht
                MyMessenger.ConsumeMessage(MyMessenger.Config.StatusQ, CType(de.Key, GunMessage).MessageID)
            Next

            '--- write a status message to say that we have reloaded
            MyMessenger.Log(Messenger.LogType.StatusType, "AppConfig reloaded at " & Format(Now, "dd-MMM-yy HH:mm:ss"))

            Dim sCmdLine As String = Environment.CommandLine
            Shell(sCmdLine, AppWinStyle.NormalNoFocus)
            End
        End If

    End Sub


    Private Sub CheckForJobDefReload()

        '--- only for Watcher
        If sTHIS_APP = "watcher" Then

            'Construct msg selector
            Dim MsgSelector As String = Replace(_MessageSelector, "#[StatusType]#", "jobdefreload")

            'If any messages returned, re-set the app config
            Dim ht As Hashtable
            ht = MyMessenger.SnapshotQueue(MyMessenger.Config.StatusQ, MsgSelector)

            '--- run it again and kill it off
            If ht.Count > 0 Then

                '--- consume the message so it doesn't trigger this again
                For Each de As DictionaryEntry In ht
                    MyMessenger.ConsumeMessage(MyMessenger.Config.StatusQ, CType(de.Key, GunMessage).MessageID)
                Next

                '--- write a status message to say that we have reloaded
                MyMessenger.Log(Messenger.LogType.StatusType, "JobDef reloaded at " & Format(Now, "dd-MMM-yy HH:mm:ss"))

                Dim sCmdLine As String = Environment.CommandLine
                Shell(sCmdLine, AppWinStyle.NormalNoFocus)
                End
            End If

        End If

    End Sub



    Private Function GetBasicFields() As Hashtable

        Dim nvp As Hashtable = New Hashtable

        With nvp
            .Add("process", sTHIS_APP)
            .Add("machine", MachineName())
            .Add("pid", CurrentPID)

            '--- other diagnostic items
            '--- e.g. appconfig file datetime, jobdef datetime


        End With
        Return nvp

    End Function



#End Region



#Region "Public"

    Public Event JobDefReloadRequired As EventHandler

    Public Enum MessageType
        Harakiri
        WhoIsOutThere
        AppConfigReload
        jobdefreload
    End Enum

    ''' <summary>
    ''' Set up module-level variables, and default (parameterised) message selector
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        'Message selector will always be the same except for the status type, so parameterise this
        _MessageSelector += "type=" & Qt("#[StatusType]#") & " "
        _MessageSelector += "AND (("
        _MessageSelector += "(machine=" & Qt(MachineName()) & " "
        _MessageSelector += "OR machine=" & Qt("all") & ") "
        _MessageSelector += "AND "
        _MessageSelector += "(process=" & Qt(sTHIS_APP) & " "
        _MessageSelector += "OR process=" & Qt("all") & ")"
        _MessageSelector += ")"
        _MessageSelector += " OR "
        _MessageSelector += "(machine=" & Qt(MachineName()) & " "
        _MessageSelector += "AND "
        _MessageSelector += "pid=" & CurrentPID() & ")"
        _MessageSelector += ")"

        '*** TESTING ***
        'Dim myNVP As New Hashtable
        'myNVP.Add("status message msg selector", _MessageSelector)
        'MyMessenger.PostMessage(MyMessenger.Config.StatusQ, "", "", "", myNVP)

    End Sub

    ''' <summary>
    ''' Define the scope of the status message, but not send it
    ''' </summary>
    ''' <param name="StatusType"></param>
    ''' <param name="Machines"></param>
    ''' <param name="Processes"></param>
    ''' <remarks></remarks>
    Public Sub SetupMessage(ByVal StatusType As String, _
                                        Optional ByVal Machines As String = "all", _
                                        Optional ByVal Processes As String = "all", _
                                        Optional ByVal PID As Integer = -1)

        'Write the status data to a hashtable
        _nvp = New Hashtable
        _nvp.Add("type", StatusType)

        If Machines.ToLower = "all" Then
            _nvp.Add("machine", "all")
        Else
            'allow comma or semi-colon separated
            Machines = Replace(Machines, ";", ",")
            _nvp.Add("machine", Machines.ToUpper)
        End If

        If Processes.ToLower = "all" Then
            _nvp.Add("process", "all")
        Else
            'allow comma or semi-colon separated
            Machines = Replace(Machines, ";", ",")
            _nvp.Add("process", Processes)
        End If

        If PID <> -1 Then
            _nvp.Add("pid", PID)
        End If

    End Sub

    ''' <summary>
    ''' Raises the status message set up in SetupMessage
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Raise()

        If _nvp Is Nothing Then
            Dim ex As New Exception("More info required to raise this status message.  Call SetupMessage first")
            Throw ex
            Exit Sub
        End If

        With MyMessenger
            _MessageID = .PostMessage(MyMessenger.Config.StatusQ, "", "", "", _nvp)
        End With

    End Sub

    ''' <summary>
    ''' Clears just the message that was raised
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()

        With MyMessenger
            .ConsumeMessage(MyMessenger.Config.StatusQ, _MessageID)
        End With

    End Sub


    ''' <summary>
    ''' Check that TIB is up - keep trying to reconnect until it is.  Then run status checks - look for a hari kiri (instruction for the process
    ''' to kill itself); look for a who is out there (report it to HQ); check for an app config reload imperrative (to pick up changed settings);
    ''' check for a job def reload imperrative (Watcher only).
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Check()

        'Check that TIB is up - keep trying to reconnect until it is
        If Not MyMessenger.TibIsUp Then
            Do
                MyMessenger = SetUpForMessaging(GetConfigFileName())
                Threading.Thread.Sleep(30000)
            Loop Until MyMessenger.TibIsUp
        End If

        'Now run the status checks
        CheckForHariKiri()
        CheckForWhoIsOutThere()
        CheckForAppConfigReload()
        CheckForJobDefReload()

    End Sub

#End Region

End Class
