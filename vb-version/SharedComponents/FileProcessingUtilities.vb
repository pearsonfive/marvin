
Imports System.IO

Module UtilitiesModule

    ''' <summary>
    ''' Sort a collection ascending or descending by the length of the item string.
    ''' </summary>
    ''' <param name="CollFiles">A collection of files</param>
    ''' <param name="SortAsc">boolean - sort ascending or not</param>
    ''' <returns>Sorted string array</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Function SortByStringLength(ByVal CollFiles() As FileInfo, ByVal SortAsc As Boolean) As String()

        Dim i As Integer, iNumChanges As Integer = 1
        Dim vTemp(CollFiles.Length - 1) As String, sTemp As String

        'Write the file collection into an array so we can swap the values around
        For i = 0 To CollFiles.Length - 1
            vTemp(i) = CollFiles(i).Name
        Next i

        'Keep looping through the array switching adjacent strings until a pass is made with no switches
        Do Until iNumChanges = 0
            iNumChanges = 0
            For i = LBound(vTemp) To UBound(vTemp) - 1
                If Len(vTemp(i)) >= Len(vTemp(i + 1)) = SortAsc Then
                    sTemp = vTemp(i)
                    vTemp(i) = vTemp(i + 1)
                    vTemp(i + 1) = sTemp
                    iNumChanges = iNumChanges + 1
                End If
            Next i
        Loop

        'Return the sorted array
        SortByStringLength = vTemp

    End Function

    ''' <summary>
    ''' Facility to unwind recognised tokens from a string.  Only "DateString,FormatString" recognised
    ''' Optional nvp hashtable allows other token/token value pairs to be passed
    ''' </summary>
    ''' <param name="TokenisedString">String including tokens</param>
    ''' <returns>String with tokens replaced with values</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Function DeTokenise(ByVal TokenisedString As String) As String

        Dim sDeTokenisedString As String = TokenisedString, sToken As String = "", sReplace As String = ""
        Dim iStartPos As Integer, iEndPos As Integer
        'Find the start of the first token
        iStartPos = InStr(1, sDeTokenisedString, "#[", CompareMethod.Text)
        'Keep looking until there are no more "token openers"
        Do Until iStartPos = 0

            'Find end of token
            iEndPos = InStr(iStartPos + 1, sDeTokenisedString, "]#", CompareMethod.Text)
            If iEndPos > 0 Then
                'only token recognised is a date string, comma, format string.  Look for the comma
                sToken = Mid(sDeTokenisedString, iStartPos + 2, iEndPos - iStartPos - 2)
                Dim iComma As Integer
                iComma = InStr(1, sToken, ",", CompareMethod.Text)
                'If one was found, separate into date and format strings
                If iComma > 0 Then
                    Dim sDate As String, sFormat As String, dDate As Date
                    sDate = Left(sToken, iComma - 1)
                    sFormat = Mid(sToken, iComma + 1)

                    Try
                        'Get a formatted string for this date
                        dDate = FindDay(sDate)
                        sReplace = Format(dDate, sFormat)
                    Catch
                        sReplace = ""
                    End Try
                Else
                    sReplace = ""
                End If
                'Replace the token with the new string
                sDeTokenisedString = Replace(sDeTokenisedString, "#[" & sToken & "]#", sReplace)

            End If

            'Find start of next token
            iStartPos = InStr(iStartPos + 1, sDeTokenisedString, "#[", CompareMethod.Text)

        Loop

        Return sDeTokenisedString

    End Function

    ''' <summary>
    '''Turn a recognised day expression into the actual date it represents.  Account for weekends and UK 
    ''' holidays.  
    ''' </summary>
    ''' <param name="DayExpression">"TODAY", "TODAY+1" or "TODAY-1"</param>
    ''' <returns>An adjusted date</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>

    Function FindDay(ByVal DayExpression As String) As Date

        Dim dDayToday As Date = Now.Date, sMessage As String

        'Find the holiday file as specified in the app config
        Dim sHols As String = ""
        For Each de As DictionaryEntry In MyMessenger.Config.HolidayFiles
            If de.Key.ToString = "GBP" Then
                sHols = de.Value.ToString()
                Exit For
            End If
        Next

        'if no hols file was found, report only but continue processing
        If sHols = "" Then
            sMessage = "#No GBP holiday file found in config file"
            MyMessenger.Log(Messenger.LogType.ErrorType, sMessage, "")
        End If

        'If couldn't read hols file, report and contine
        Try
            sHols = IO.File.OpenText(sHols).ReadToEnd
        Catch ex As Exception
            sMessage = "#Unable to open GBP holiday file: " & sHols
            MyMessenger.Log(Messenger.LogType.ErrorType, sMessage, "")
        End Try

        'Adjust the day according to the holiday file (if any)
        Select Case DayExpression

            Case "TODAY-1"
                'Adjust for weekend
                If Weekday(Now) = 2 Then
                    FindDay = dDayToday.AddDays(-3)
                Else
                    FindDay = dDayToday.AddDays(-1)
                End If
                'Keep moving the date backwards until find one that's not in the holiday file
                Do Until InStr(1, sHols, FindDay.ToString("dd-MMM-yyyy"), CompareMethod.Text) = 0
                    FindDay = dDayToday.AddDays(-1)
                Loop

            Case "TODAY"
                FindDay = dDayToday

            Case "TODAY+1"
                If Weekday(Now) = 6 Then
                    FindDay = dDayToday.AddDays(3)
                Else
                    FindDay = dDayToday.AddDays(1)
                End If
                'Keep moving the date forwards until find one that's not in the holiday file
                Do Until InStr(1, sHols, FindDay.ToString("dd-MMM-yyyy"), CompareMethod.Text) = 0
                    FindDay = dDayToday.AddDays(1)
                Loop

            Case Else
                Err.Raise(-15031967, "FindDay", "Unable to determine date")

        End Select

    End Function

    ''' <summary>
    ''' Shell out a command line to a process, provision to send switches and allow a timeout.  See
    ''' http://www.developerfusion.co.uk/show/4662/
    ''' </summary>
    ''' <param name="CommandLine">Path to the exe</param>
    ''' <param name="Parameters">String of switches</param>
    ''' <param name="TimeOutSeconds">Timeout in seconds</param>
    ''' <returns>Whatever the app writes to the pipe</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>

    Function ShellAndWait(ByVal CommandLine As String, _
                                    ByVal Parameters As String, _
                                    Optional ByVal TimeOutSeconds As Integer = 0, _
                                    Optional ByVal WindowStyle As ProcessWindowStyle = ProcessWindowStyle.Hidden) As String

        'Run the command line.  
        Dim proc As Process = New Process
        With proc

            .StartInfo.FileName = CommandLine
            .StartInfo.Arguments = Parameters
            .StartInfo.UseShellExecute = False            ' need to set this to false to redirect output
            .StartInfo.RedirectStandardOutput = True
            .Start()
            Dim fHasRunProperly As Boolean = .WaitForExit(TimeOutSeconds * 1000)
            If Not fHasRunProperly Then
                .Kill()
                Return ""
            Else
                Return .StandardOutput.ReadToEnd
            End If

        End With

    End Function


    Public Function CheckFilePath(ByVal FilePathToTest As String) As Boolean

        If FilePathToTest <> "" Then
            If My.Computer.FileSystem.FileExists(FilePathToTest) Then Return True
            If My.Computer.FileSystem.DirectoryExists(FilePathToTest) Then Return True
        End If

        Return False

    End Function



    Public Function GetGoodFileName(ByVal FilePath As String, _
                                                    ByVal FileName As String, _
                                                    Optional ByVal UseSuppliedName As Boolean = True) As String

        Dim OriginalFileName As String = Path.Combine(FilePath, FileName)

        If Not File.Exists(OriginalFileName) Then
            'Filename doesn't already exist, so will use this name
            Return OriginalFileName

        Else

            'Parse into root file name and extension
            Dim RootName As String = Left(FileName, InStrRev(FileName, ".") - 1)
            Dim FileExtension As String = Mid(FileName, InStrRev(FileName, "."))

            'iterate until get to a filename that doesn't exist
            Dim i As Integer = 0
            Do
                i += 1
            Loop Until Not File.Exists(Path.Combine(FilePath, RootName & FileExtension & "." & i.ToString))
            Dim NewFileName As String = Path.Combine(FilePath, RootName & FileExtension & "." & i.ToString)

            If UseSuppliedName Then
                'Rename the origiinal file to have the new name
                File.Move(Path.Combine(FilePath, FileName), NewFileName)
                'Return the original filename, which now doesn't exist
                Return Path.Combine(FilePath, FileName)
            Else
                'Return the new (unused) name
                Return NewFileName
            End If
        End If

    End Function

End Module
