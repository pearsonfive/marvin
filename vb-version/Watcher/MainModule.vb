Imports System.IO
Imports System.Net.Mail



Module MainModule


    ''' <summary>
    ''' Entry point for app
    ''' Command-line args are used to setup the app
    ''' </summary>
    ''' <remarks>SDCP 13-Sep-2006</remarks>
    Sub Main()

        Try

            Console.WriteLine("This is Watcher")
            Console.WriteLine()
            Console.WriteLine(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe.")
            Console.WriteLine()
            Console.Write("Built: " & Format(FileSystem.FileDateTime(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe."), "dd-MMM-yyyy HH:mm:ss"))
            Console.WriteLine()
            Console.WriteLine(Date.Now.ToLongTimeString & " - " & "Start up")
            Console.WriteLine()

            '--- set up the Messenger
            Prerequisites()

            '--- check status messages
            MyStatusMessages.Check()


            '--- do the needful w.r.t. setting up watchers etc
            Console.WriteLine(Date.Now.ToLongTimeString & " - " & "StartProcessing()")
            StartProcessing()
            Console.WriteLine(Date.Now.ToLongTimeString & " - " & "StartProcessing() ended")
            Console.WriteLine()
            Console.WriteLine("Running...")


            'Err.Raise(vbObjectError + 1001, "Deliberate Error")


            '--- loop looking for the Hara-Kiri package
            Do

                '--- check status messages
                MyStatusMessages.Check()

                'Console.WriteLine(Date.Now.ToLongTimeString & " - " & "About to sleep for " & MyMessenger.Config.PollFrequency & " seconds.")

                '--- sleep the thread for some time
                Threading.Thread.Sleep(MyMessenger.Config.PollFrequency * 1000)

            Loop

        Catch ex As Exception

            'Spawn a new instance of Watcher
            Dim sWatcherExeLocation As String = Environment.GetCommandLineArgs(0)
            Shell(sWatcherExeLocation & " " & MySwitches.FullCommandLine, AppWinStyle.NormalNoFocus, False)

            '--- create the attachments collection
            Dim atts As New Collection
            Try
                Dim sConfigFilename As String = Path.Combine(Environ("temp"), "Config_" & Now.ToString("yyyyMMdd_HHmmss") & ".xml")
                MyMessenger.Config.DomDoc.Save(sConfigFilename)
                atts.Add(New Attachment(sConfigFilename))
            Catch
            End Try
            Try
                Dim sJobDefFilename As String = Path.Combine(Environ("temp"), "JobDef_" & Now.ToString("yyyyMMdd_HHmmss") & ".xml")
                Parser.JobDefinitionsXml.Save(sJobDefFilename)
                atts.Add(New Attachment(sJobDefFilename))
            Catch
            End Try
            Try
                MyMessenger.SendEmail("DL-RatesDeskDev@ubs.com", "DL-RatesDeskDev@ubs.com", "Watcher failure on " & My.Computer.Name, ex.Message & vbCrLf & vbCrLf & ex.StackTrace & HarvestProcessingDetails("commandline|switches"), atts)
                MyMessenger.Log(Messenger.LogType.ErrorType, ex.StackTrace)

            Catch
            End Try

            LastDitchLog(ex.StackTrace)

        End Try

    End Sub


    ''' <summary>
    ''' Items that must be run at startup
    ''' </summary>
    ''' <remarks>SDCP 18-Oct-2006</remarks>
    Private Sub Prerequisites()

        '--- command-line switches
        MySwitches = New Switches

        '--- create a messenger
        MyMessenger = SetUpForMessaging(GetConfigFileName)

        '--- create a status message handler
        MyStatusMessages = New StatusMessage()

    End Sub


    Public Function GetWatcherErrorNVPs() As Hashtable
        Dim ht As Hashtable

        ht = New Hashtable

        '--- build the NVPs
        ht.Add("type", "WatcherError")
        ht.Add("computer", My.Computer.Name)
        ht.Add("pid", Process.GetCurrentProcess.Id)

        Return ht

    End Function


End Module
