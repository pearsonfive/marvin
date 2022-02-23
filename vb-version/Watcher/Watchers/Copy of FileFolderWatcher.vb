Imports System.IO

Public Class FileFolderWatcher : Inherits WatcherBase

    Private _mask As String
    Private _folder As String
    Private _lookInSubfolders As Boolean
    '
    Private _filesFoundUpToLastTime As Hashtable
    Private _combinationLogic As String
    Private _watcherInfoList As List(Of WatcherInfo)

    Private Const NOT_WRITTEN_TO_FOR_THIS_MANY_SECONDS As Integer = 60


    ''' <summary>
    ''' Constructor for FileFolderWatcher
    ''' </summary>
    ''' <param name="ThisJob">The job object that we are going to watch for</param>
    ''' <remarks>SDCP 10-Oct-2006</remarks>
    Sub New(ByVal ThisJob As Job)

        '--- store the active trigger
        _activeJob = ThisJob

        '--- instantiate the hashtables
        _filesFoundUpToLastTime = New Hashtable

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

        '--- de-serialise the info from the xml
        For Each ndWatcher As Xml.XmlNode In _activeJob.Trigger.TriggerInfoXml.SelectNodes("descendant::watcher")

            '--- create a SetupInfo for each watcher
            Dim aWatcherInfo As New WatcherInfo

            With ndWatcher
                aWatcherInfo.Folder = .SelectSingleNode("folder").InnerText
                aWatcherInfo.Mask = .SelectSingleNode("mask").InnerText
                aWatcherInfo.LookInSubFolders = CBool(.SelectSingleNode("folder/@includesubfolders").InnerText)
                aWatcherInfo.MaxAgeUnit = .SelectSingleNode("maxtriggerage/@unit").InnerText
                aWatcherInfo.MaxAgeAmount = CInt(.SelectSingleNode("maxtriggerage/@amount").InnerText)
                Try
                    aWatcherInfo.LastWrittenAgeUnit = .SelectSingleNode("lastwrittenage/@unit").InnerText
                Catch ex As Exception
                    aWatcherInfo.LastWrittenAgeUnit = "second"
                End Try
                Try
                    aWatcherInfo.LastWrittenAgeAmount = CInt(.SelectSingleNode("lastwrittenage/@amount").InnerText)
                Catch ex As Exception
                    aWatcherInfo.LastWrittenAgeAmount = 60
                End Try
            End With

            '--- add to the list
            _watcherInfoList.Add(aWatcherInfo)

        Next

    End Sub


    ''' <summary>
    ''' If we want to stop the watcher
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 10-Oct-2006</remarks>
    Public Property StopWatcher() As Boolean
        Get
            Return _stopWatcher
        End Get
        Set(ByVal value As Boolean)
            _stopWatcher = value
        End Set
    End Property


    ''' <summary>
    ''' property handler for Mask
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 06-Sep-2006</remarks>
    Public ReadOnly Property Mask() As String
        Get
            Mask = _mask
        End Get
    End Property


    ''' <summary>
    ''' property handler for Folder
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 06-Sep-2006</remarks>
    Public ReadOnly Property Folder() As String
        Get
            Folder = _folder
        End Get
    End Property


    ''' <summary>
    ''' property handler for LookInSubFolders
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 06-Sep-2006</remarks>
    Public ReadOnly Property LookInSubFolders() As Boolean
        Get
            LookInSubFolders = _lookInSubfolders
        End Get
    End Property



    ''' <summary>
    ''' polls the file/folder for changes since last poll time
    ''' </summary>
    ''' <remarks>SDCP 06-Sep-2006</remarks>
    Protected Overrides Sub Poller()
        Dim di As DirectoryInfo
        Dim subdi As DirectoryInfo
        Dim fi As FileInfo
        Dim so As SearchOption = SearchOption.TopDirectoryOnly
        Dim subdirectoryMask As String
        Dim baseFolder As String


        Try


            '--- poll for each watcher
            For Each aWatcherInfo As WatcherInfo In _watcherInfoList

                With aWatcherInfo

                    '.Folder = "hello world"

                    '_watcherInfoList(0).Folder = "hello arse"

                    '--- create a hashtable for the found triggers
                    .TriggersFound = New Hashtable
                    .FilesFoundUpToLastTime = New Hashtable

                    '--- set SearchOption
                    If .LookInSubFolders Then so = SearchOption.AllDirectories

                    Dim iWildcardPos As Integer = InStr(1, .Folder, "*", CompareMethod.Text)

                    If iWildcardPos > 0 Then
                        subdirectoryMask = .Folder.Substring(iWildcardPos - 1)
                        baseFolder = .Folder.Substring(0, iWildcardPos - 1)
                    Else
                        subdirectoryMask = ""
                        baseFolder = .Folder
                    End If

                    '--- get the part of the path that doesn't have a wildcard - check it exists
                    di = New DirectoryInfo(baseFolder)

                    'ML - Watcher crashing if the folder's not found (eg drive map missing, or typo in JobDef)
                    If Not FileIO.FileSystem.DirectoryExists(baseFolder) Then
                        Dim ErrMessage As String = "Watcher Error - Invalid directory path in JobDef: "
                        ErrMessage += vbCr & baseFolder
                        MyMessenger.Log(Messenger.LogType.ErrorType, ErrMessage)
                        Exit Sub
                    End If

                    '--- calculate the earliest Last Modified date we are interested in
                    '--- (this is so we can throttle the number of triggers that we check)
                    Dim earliestLastModified As Date = GetAdjustedDate(Now, .MaxAgeUnit, .MaxAgeAmount * -1)
                    Dim latestLastWrittenTo As Date = GetAdjustedDate(Now, .LastWrittenAgeUnit, .LastWrittenAgeAmount * -1)


                    '--- get the fileinfos in the base folder
                    If subdirectoryMask = "" Then

                        '--- get the FileInfos for all files matching the mask
                        For Each fi In di.GetFiles(.Mask, so)

                            '--- only if younger than max trigger age
                            If fi.LastWriteTime > earliestLastModified Then

                                '--- RECENTLY MODIFIED FILES first
                                '--- only if modified since last poll time
                                If fi.LastWriteTime >= _lastPollDateTime Then

                                    '--- only if the trigger file has not been written to recently
                                    '--- (i.e. may still be being written to)
                                    If fi.LastWriteTime < latestLastWrittenTo Then

                                        '--- add to the triggers found hashtable
                                        .TriggersFound.Add(fi.FullName, fi)

                                        '--- add to the cumulative files found hashtable if not there already
                                        If Not (.FilesFoundUpToLastTime.ContainsKey(fi.FullName)) Then
                                            .FilesFoundUpToLastTime.Add(fi.FullName, fi)
                                        End If

                                    End If

                                End If

                                ''--- NEW FILES, e.g. copied in from another folder
                                ''--- but only if not already in the cumulative list
                                'If Not (.FilesFoundUpToLastTime.ContainsKey(fi.FullName)) Then

                                '    '--- add to the triggers found hashtable
                                '    .TriggersFound.Add(fi.FullName, fi)

                                '    '--- add to the cumulative files found hashtable
                                '    .FilesFoundUpToLastTime.Add(fi.FullName, fi)

                                'End If

                            End If

                        Next

                        ''--- NEW FILES, e.g. copied in from another folder
                        ''--- but only if not already in the cumulative list
                        'For Each fi In di.GetFiles(.Mask, so)

                        '    If Not (.FilesFoundUpToLastTime.ContainsKey(fi.FullName)) Then

                        '        '--- add to the triggers found hashtable
                        '        .TriggersFound.Add(fi.FullName, fi)

                        '        '--- add to the cumulative files found hashtable
                        '        .FilesFoundUpToLastTime.Add(fi.FullName, fi)

                        '    End If

                        'Next fi



                    Else '--- if there is a mask for subirectories then we do not look in the base folder

                        '--- loop through all subdirectories
                        For Each subdi In di.GetDirectories(subdirectoryMask)

                            '--- get the FileInfos for all files matching the mask
                            For Each fi In subdi.GetFiles(.Mask, so)

                                '--- only if younger than max trigger age
                                If fi.LastWriteTime > earliestLastModified Then

                                    '--- RECENTLY MODIFIED FILES first
                                    '--- only if modified since last poll time
                                    If fi.LastWriteTime >= _lastPollDateTime Then

                                        '--- only if the trigger file has not been written to recently
                                        '--- (i.e. may still be being written to)
                                        If fi.LastWriteTime < latestLastWrittenTo Then

                                            '--- add to the triggers found hashtable
                                            aWatcherInfo.TriggersFound.Add(fi.FullName, fi)

                                            '--- add to the cumulative files found hashtable if not there already
                                            If Not (.FilesFoundUpToLastTime.ContainsKey(fi.FullName)) Then
                                                .FilesFoundUpToLastTime.Add(fi.FullName, fi)
                                            End If

                                        End If

                                    End If

                                    ''--- NEW FILES, e.g. copied in from another folder
                                    ''--- but only if not already in the cumulative list
                                    'If Not (.FilesFoundUpToLastTime.ContainsKey(fi.FullName)) Then

                                    '    '--- add to the triggers found hashtable
                                    '    aWatcherInfo.TriggersFound.Add(fi.FullName, fi)

                                    '    '--- add to the cumulative files found hashtable
                                    '    .FilesFoundUpToLastTime.Add(fi.FullName, fi)

                                    'End If

                                End If

                            Next fi

                        Next subdi

                        ' ''--- NEW FILES, e.g. copied in from another folder
                        ' ''--- but only if not already in the cumulative list
                        ''--- loop through all subdirectories
                        'For Each subdi In di.GetDirectories(subdirectoryMask)

                        '    '--- get the FileInfos for all files matching the mask
                        '    For Each fi In subdi.GetFiles(.Mask, so)

                        '        If Not (.FilesFoundUpToLastTime.ContainsKey(fi.FullName)) Then

                        '            '--- add to the triggers found hashtable
                        '            aWatcherInfo.TriggersFound.Add(fi.FullName, fi)

                        '            '--- add to the cumulative files found hashtable
                        '            .FilesFoundUpToLastTime.Add(fi.FullName, fi)

                        '        End If

                        '    Next fi

                        'Next subdi


                    End If

                End With

            Next



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


            If FoundValidTrigger = True Then

                '--- populate the TriggersFound table
                _triggersFound = New Hashtable
                For Each aWatcherInfo As WatcherInfo In _watcherInfoList

                    For Each de As DictionaryEntry In aWatcherInfo.TriggersFound

                        _triggersFound.Add(de.Key, de.Value)

                    Next

                Next aWatcherInfo

                CreateMessages()
            End If


        Catch ex As Exception

            Select Case ex.GetType.ToString

                Case GetType(IO.FileNotFoundException).ToString, GetType(IO.DirectoryNotFoundException).ToString, GetType(IO.DriveNotFoundException).ToString
                    Dim ErrMessage As String = "FileFolderWatcher Error - File, Folder or Drive not found. "

                    If ex.GetType.ToString = GetType(IO.FileNotFoundException).ToString Then
                        ErrMessage += vbCr & "(" & CType(ex, IO.FileNotFoundException).FileName & ")"
                    End If

                    ErrMessage += vbCr & "Retrying..."
                    ErrMessage += vbCr & ex.StackTrace
                    MyMessenger.Log(Messenger.LogType.ErrorType, ErrMessage)

                Case Else

                    Dim ErrMessage As String = "General FileFolderWatcher Error"
                    ErrMessage += vbCr & "Retrying..."
                    ErrMessage += vbCr & ex.StackTrace
                    MyMessenger.Log(Messenger.LogType.ErrorType, ErrMessage)

            End Select

            Exit Sub

        End Try

    End Sub


    ''' <summary>
    ''' helper sub-class to store settings for each watcher
    ''' </summary>
    ''' <remarks></remarks>
    Class WatcherInfo : Inherits WatcherBase.WatcherInfoBase
        Public Folder As String
        Public Mask As String
        Public LookInSubFolders As Boolean
        Public FilesFoundUpToLastTime As Hashtable
        Public LastWrittenAgeUnit As String
        Public LastWrittenAgeAmount As Integer
    End Class


End Class




