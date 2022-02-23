Imports System.Threading

Public Class WatcherManager

    '--- create events
    Public Event WatcherCreationStart(ByVal sender As Object, ByVal e As EventArgs)
    Public Event WatcherCreationEnd(ByVal sender As Object, ByVal e As EventArgs)



    ''' <summary>
    ''' poll the triggers and create watchers as required
    ''' </summary>
    ''' <remarks>SDCP 26-Sep-2006</remarks>
    Public Sub CreateWatcherThreads()

        '--- iterate through jobgroups and jobs to build the triggers
        For Each jg As JobGroup In JobGroupList

            For Each j As Job In jg

                '--- only if active
                If j.JobXml.SelectSingleNode("descendant::status").InnerText = "active" Then

                    'zHACK: this is horrible
                    '--- stop threads stepping over each other
                    Threading.Thread.Sleep(100)

                    'Because Martin keeps building in single-threaded mode!
                    If InStr(1, Environment.GetCommandLineArgs(0), "vshost.exe", CompareMethod.Text) = 0 Then

                        ''--- create the watcher on a new thread
                        Dim t As New Thread(AddressOf CreateWatcher)
                        t.Start(j)

                    Else

                        'For testing, run non-threaded, single run
                        CreateWatcher(j)

                    End If

                End If

            Next

        Next


    End Sub


    ''' <summary>
    ''' setup the watchers based on the the incoming Job object
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks>SDCP 26-Sep-2006</remarks>
    Private Sub CreateWatcher(ByVal obj As Object)

        Dim ThisJob As Job = CType(obj, Job)

        '--- different setup methodology for each Watcher
        Select Case ThisJob.Trigger.WatcherType

            Case "FileFolderWatcher"

                Dim ffw As FileFolderWatcher = New FileFolderWatcher(ThisJob)

            Case "QueueWatcher"

                Dim qw As QueueWatcher = New QueueWatcher(ThisJob)

            Case "TopicWatcher"

                Dim tw As TopicWatcher = New TopicWatcher(ThisJob)

            Case "PollingWatcher"

                Dim pw As PollingWatcher = New PollingWatcher(ThisJob)


            Case Else

                '*** testing ***


        End Select

    End Sub


End Class
