Imports System.IO
Imports OTIS
Imports System.Collections.Specialized


Public Class SymbolicsExpander


    '--- local variables
    Private _anyText As String
    Private _triggerXml As Xml.XmlNode
    Private _watcher As WatcherBase
    Private _expandedString As String

    Private _expandedSymbols As ArrayList
    Private _unExpandedSymbols As ArrayList

    Private _pulledTrigger As Trigger
    Private _fileInfoTrigger As FileInfo
    Private _gunMessageTrigger As GunMessage

    Private _allSymbolsFound As Hashtable
    'Private _source As String
    Private _unExpandedString As String

    Private _fixedNow As Date = Now()
    Private _fixedUniqueID As String = ""

    Private _scope As Scope

    Private Const LHS As String = "#["
    Private Const RHS As String = "]#"


    Public Enum Scope
        PreTrigger
        PostTrigger
    End Enum


    '--- constructors
    'Sub New(ByVal AnyText As String)
    '    _source = "AnyText"
    '    _anyText = AnyText
    'End Sub

    Sub New(ByVal TriggerXml As Xml.XmlNode, ByVal Scope As Scope)

        '_source = "PreTriggering"
        _scope = Scope

        '--- set up local variables
        _triggerXml = TriggerXml

    End Sub

    Sub New(ByVal TriggerXml As Xml.XmlNode, ByVal Watcher As WatcherBase, ByVal PulledTrigger As Object, ByVal Scope As Scope)

        '_source = "PostTriggering"
        _scope = Scope


        '--- set up local variables
        _triggerXml = TriggerXml
        _watcher = Watcher

        '--- define values for Tokens that are fixed values for this trigger message
        '--- e.g. #[NOW]# and #[UNIQUE_ID]#
        _fixedUniqueID = CreateUniqueID()
        _fixedNow = Now()

        '--- get the trigger object
        _pulledTrigger = Watcher.ActiveJob.Trigger

        '--- cast trigger object to required type for watcher
        Select Case True

            Case TypeOf Watcher Is FileFolderWatcher
                _fileInfoTrigger = CType(PulledTrigger, FileInfo)

            Case TypeOf Watcher Is QueueWatcher
                _gunMessageTrigger = CType(PulledTrigger, GunMessage)

            Case Else
                '
                '
        End Select


    End Sub


    ''' <summary>
    ''' expands 1 symbol
    ''' </summary>
    ''' <param name="Symbol"></param>
    ''' <param name="SymbolAndFormat"></param>
    ''' <returns></returns>
    ''' <remarks>SDCP 07-Sep-2007</remarks>
    Private Function ExpandIndividualSymbol(ByVal Symbol As String, ByVal SymbolAndFormat As String) As String

        Dim sExpandedSymbol As String = LHS & SymbolAndFormat & RHS


        '--- check that the symbol  is allowed by the NonContextualSymbols hashtable
        If MyMessenger.Config.NonContextualSymbols(Symbol) IsNot Nothing Then

            Select Case Symbol.ToUpper()

                Case "TODAY", "TODAY-1", "TODAY+1"
                    sExpandedSymbol = ExpandToday(SymbolAndFormat)

                Case "NOW"
                    sExpandedSymbol = ExpandNow(SymbolAndFormat)

                Case "TRIGGER_TODAY", "TRIGGER_TODAY-1", "TRIGGER_TODAY+1"
                    sExpandedSymbol = ExpandTriggerToday(SymbolAndFormat)

                Case "TRIGGER_NOW"
                    sExpandedSymbol = ExpandTriggerNow(SymbolAndFormat)

                Case "TRIGGER_FILE_NAME"
                    sExpandedSymbol = ExpandTriggerFilename(SymbolAndFormat)
                    '    If sExpandedSymbol <> (LHS & SymbolAndFormat & RHS) Then sExpandedSymbol = Path.GetFileName(sExpandedSymbol)

                Case "TRIGGER_FILE_NAME_NO_EXTENSION"
                    sExpandedSymbol = ExpandTriggerFilenameNoExtension(SymbolAndFormat)
                    '   If sExpandedSymbol <> (LHS & SymbolAndFormat & RHS) Then sExpandedSymbol = Path.GetFileNameWithoutExtension(sExpandedSymbol)

                Case "TRIGGER_FILE_PATH"
                    sExpandedSymbol = ExpandTriggerFilePath(SymbolAndFormat)
                    '  If sExpandedSymbol <> (LHS & SymbolAndFormat & RHS) Then sExpandedSymbol = Path.GetDirectoryName(sExpandedSymbol)

                Case "TRIGGER_FILE_FULLPATH"
                    sExpandedSymbol = ExpandTriggerFileFullPath(SymbolAndFormat)

                Case "TRIGGER_FILE_LASTMODIFIED"
                    sExpandedSymbol = ExpandTriggerFileLastModified(SymbolAndFormat)

                Case "TRIGGER_MESSAGE_ID"
                    sExpandedSymbol = ExpandTriggerMessageId(SymbolAndFormat)

                Case "TRIGGER_MESSAGE_NVP"
                    sExpandedSymbol = ExpandTriggerMessageNvp(SymbolAndFormat)

                Case "TRIGGER_MESSAGE_PAYLOAD"
                    sExpandedSymbol = ExpandTriggerMessagePayload(SymbolAndFormat)

                Case "UNIQUE_ID"
                    sExpandedSymbol = ExpandUniqueId(SymbolAndFormat)

                Case "JOB_NAME"
                    sExpandedSymbol = ExpandJobName(SymbolAndFormat)

                Case "QUEUE_ALIAS"
                    sExpandedSymbol = ExpandQueueAlias(SymbolAndFormat)

                    '--- fixed items
                Case "TRIGGER_UNIQUE_ID"
                    sExpandedSymbol = ExpandTriggerUniqueId(SymbolAndFormat)

                Case "TRIGGER_TIMESTAMP"
                    sExpandedSymbol = ExpandTriggerTimestamp(SymbolAndFormat)

                Case "SUBSTRING"
                    sExpandedSymbol = ExpandSubstring(SymbolAndFormat)

                Case "DATEVALUE"
                    sExpandedSymbol = ExpandDateValue(SymbolAndFormat)

                Case "TIMEVALUE"
                    sExpandedSymbol = ExpandTimeValue(SymbolAndFormat)

                Case "ENVIRON"
                    sExpandedSymbol = ExpandEnviron(SymbolAndFormat)

                Case "ROUNDTIME"
                    sExpandedSymbol = ExpandRoundTime(SymbolAndFormat)

                Case Else
                    sExpandedSymbol = LHS & SymbolAndFormat & RHS

            End Select

            '--- if symbol has been expanded then need to escape entity references
            If sExpandedSymbol <> LHS & SymbolAndFormat & RHS Then
                sExpandedSymbol = Escape(sExpandedSymbol)
            End If

        End If

        Return sExpandedSymbol

    End Function


    ''' <summary>
    ''' Entry point for expansion
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Expand()

        '--- extract the unexpanded string from source
        Select Case _scope

            'Case "AnyText"
            '    _unExpandedString = _anyText

            Case Scope.PreTrigger
                _unExpandedString = _triggerXml.OuterXml

            Case Scope.PostTrigger
                _unExpandedString = _triggerXml.OuterXml

        End Select

        If _scope = Scope.PreTrigger Then

            '--- expand the contextual symbols
            _unExpandedString = ExpandContextualSymbols()

        End If

        'Else

        '--- the remaining symbols are non-contextual

        '--- get the collection of symbols  in the string
        _allSymbolsFound = GetAllSymbols()

        '--- expand symbols in the hashtable
        'ExpandNonContextualSymbols()
        _expandedString = ExpandNonContextualSymbols_TheNewWay(_unExpandedString)

        '--- replace each instance of the symbol in the unexpanded string
        'ReplaceWithExpansions()

        'End If

    End Sub


    ''' <summary>
    ''' Creates a hashtable of all the symbols (valid or otherwise) in the text
    ''' </summary>
    ''' <remarks></remarks>
    Private Function ExpandNonContextualSymbols_TheNewWay(ByVal ExpandThis As String) As String
        Dim lStartPos As Integer
        Dim lNextStartPos As Integer
        Dim lEndPos As Integer
        Dim sSymbol As String
        Dim sExpandedSymbol As String

        Dim sImprobableLHS As String = "LHS_" & CreateUniqueID()
        Dim sImprobableRHS As String = "RHS_" & CreateUniqueID()

        lNextStartPos = 1

        Try


            Do

                '--- find first LHS
                lStartPos = InStr(Math.Max(lNextStartPos, 1), ExpandThis, LHS, CompareMethod.Text)

                '--- find next RHS
                lEndPos = InStr(lStartPos + 1, ExpandThis, RHS, CompareMethod.Text)

                '--- exit if required
                If lStartPos = 0 Then

                    '--- if no RHS either then exit
                    If lEndPos = 0 Then
                        Exit Do
                    Else
                        '--- back to the beginning for 1 more pass
                        lStartPos = InStr(1, ExpandThis, LHS, CompareMethod.Text)
                        lEndPos = InStr(lStartPos + 1, ExpandThis, RHS, CompareMethod.Text)
                    End If

                End If

                '--- and find next LHS
                lNextStartPos = InStr(lStartPos + 1, ExpandThis, LHS, CompareMethod.Text)

                '--- if RHS is first then we know that we have something to expand
                If (lEndPos < lNextStartPos) Or (lNextStartPos = 0) Then

                    '--- extract the Symbol
                    sSymbol = Mid$(ExpandThis, lStartPos + 2, lEndPos - lStartPos - 2)

                    '--- and expand
                    sExpandedSymbol = ExpandIndividualSymbol(Split(sSymbol, "|")(0), sSymbol)

                    '--- unexpanded symbols need to be escaped with the improbable delimiters so that they don't get considered again
                    sExpandedSymbol = sExpandedSymbol.Replace(LHS, sImprobableLHS)
                    sExpandedSymbol = sExpandedSymbol.Replace(RHS, sImprobableRHS)

                    '--- replace in original text
                    ExpandThis = ExpandThis.Remove(lStartPos - 1, Len(sSymbol) + 4)
                    ExpandThis = ExpandThis.Insert(lStartPos - 1, sExpandedSymbol)

                End If

                '--- do it all again
                lStartPos = 0
                lEndPos = 0

            Loop




            '--- replace the "improbable" delimiters
            ExpandThis = ExpandThis.Replace(sImprobableLHS, LHS)
            ExpandThis = ExpandThis.Replace(sImprobableRHS, RHS)


        Catch ex As Exception

            MyMessenger.Log(Messenger.LogType.ErrorType, ex.StackTrace)


        End Try


        '--- expose the expanded string
        Return ExpandThis

    End Function





    ''' <summary>
    ''' Creates a hashtable of all the symbols (valid or otherwise) in the text
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAllSymbols() As Hashtable
        Dim ht As Hashtable
        Dim lStartPos As Integer
        Dim lEndPos As Integer
        Dim sSymbol As String


        ht = New Hashtable

        '--- first test
        lStartPos = InStr(1, _unExpandedString, LHS, CompareMethod.Text)

        '--- loop until no more are found
        Do Until lStartPos = 0

            '--- find the end of the token
            lEndPos = InStr(lStartPos + 1, _unExpandedString, RHS, CompareMethod.Text)

            '--- only continue if RHS found
            If lEndPos > 0 Then

                '--- extract the Symbol
                sSymbol = Mid$(_unExpandedString, lStartPos + 2, lEndPos - lStartPos - 2)

                '--- add the symbol to the name value collection
                Try
                    ht.Add(sSymbol, LHS & sSymbol & RHS)
                Catch ex As Exception
                    'Stop
                End Try

            End If

            '--- check again
            lStartPos = 0
            lStartPos = InStr(lEndPos + 1, _unExpandedString, LHS, CompareMethod.Text)

        Loop


        '--- return the array list
        Return ht

    End Function


    ''' <summary>
    ''' Expands the contextual symbols if possible
    ''' </summary>
    ''' <remarks></remarks>
    Private Function ExpandContextualSymbols() As String
        Dim sExpandedSymbol As String = ""

        '--- loop through all the "job" nodes in the Xml and check to see if the 
        '--- inner xml contains a symbol
        For Each xJob As Xml.XmlNode In _triggerXml.SelectNodes("//job")

            '--- #[JOB_GROUP_NAME]#:
            '-- get the jobgroup name
            Dim sJobGroupName As String = xJob.SelectSingleNode("ancestor::jobgroup").Attributes("name").Value

            '--- expand the symbol
            xJob.InnerXml = xJob.InnerXml.Replace("#[JOB_GROUP_NAME]#", sJobGroupName)

        Next

        '--- return the expanded string
        '_expandedString = _triggerXml.OuterXml
        Return _triggerXml.OuterXml()

    End Function


    ''' <summary>
    ''' Expands the non-contextual symbols if possible
    ''' </summary>
    ''' <remarks>Also populates the ExpandedSymbols and UnexpandedSymbols lists too</remarks>
    Private Sub ExpandNonContextualSymbols()
        Dim sSymbolAndFormat As String
        Dim sSymbol As String
        Dim sExpandedSymbol As String = ""

        '--- prepare the lists
        _expandedSymbols = New ArrayList
        _unExpandedSymbols = New ArrayList

        '--- loop through the Symbols hashtable and attempt to process each one
        '--- (the key is the symbol and format, e.g. TODAY|dd-mmm-yyyyy)
        For Each de As DictionaryEntry In CType(_allSymbolsFound.Clone, Hashtable)

            '--- get the symbol and format
            sSymbolAndFormat = CType(de.Key, String)
            sSymbol = Split(sSymbolAndFormat, "|")(0)

            '--- initialise the expanded symbol with the unexpanded one
            sExpandedSymbol = LHS & sSymbolAndFormat & RHS

            '--- check that the symbol  is allowed by the NonContextualSymbols hashtable
            If MyMessenger.Config.NonContextualSymbols(sSymbol) IsNot Nothing Then

                sExpandedSymbol = ExpandIndividualSymbol(sSymbol, sSymbolAndFormat)

                'Select Case sSymbol.ToUpper()

                '    Case "TODAY", "TODAY-1", "TODAY+1"
                '        sExpandedSymbol = ExpandToday(sSymbolAndFormat)

                '    Case "NOW"
                '        sExpandedSymbol = ExpandNow(sSymbolAndFormat)

                '    Case "TRIGGER_TODAY", "TRIGGER_TODAY-1", "TRIGGER_TODAY+1"
                '        sExpandedSymbol = ExpandToday(sSymbolAndFormat)

                '    Case "TRIGGER_NOW"
                '        sExpandedSymbol = ExpandNow(sSymbolAndFormat)

                '    Case "TRIGGER_FILE_NAME"
                '        sExpandedSymbol = Path.GetFileName(ExpandTriggerFilename(sSymbolAndFormat))

                '    Case "TRIGGER_FILE_NAME_NO_EXTENSION"
                '        sExpandedSymbol = Path.GetFileNameWithoutExtension(ExpandTriggerFilename(sSymbolAndFormat))

                '    Case "TRIGGER_FILE_PATH"
                '        sExpandedSymbol = Path.GetDirectoryName(ExpandTriggerFilename(sSymbolAndFormat))

                '    Case "TRIGGER_FILE_FULLPATH"
                '        sExpandedSymbol = ExpandTriggerFilename(sSymbolAndFormat)

                '    Case "TRIGGER_FILE_LASTMODIFIED"
                '        sExpandedSymbol = ExpandTriggerFileLastModified(sSymbolAndFormat)

                '    Case "TRIGGER_MESSAGE_ID"
                '        sExpandedSymbol = ExpandTriggerMessageId(sSymbolAndFormat)

                '    Case "TRIGGER_MESSAGE_NVP"
                '        sExpandedSymbol = ExpandTriggerMessageNvp(sSymbolAndFormat)

                '    Case "TRIGGER_MESSAGE_PAYLOAD"
                '        sExpandedSymbol = ExpandTriggerMessagePayload(sSymbolAndFormat)

                '    Case "UNIQUE_ID"
                '        sExpandedSymbol = ExpandUniqueId(sSymbolAndFormat)

                '    Case "JOB_NAME"
                '        sExpandedSymbol = ExpandJobName(sSymbolAndFormat)

                '    Case "QUEUE_ALIAS"
                '        sExpandedSymbol = ExpandQueueAlias(sSymbolAndFormat)

                '    Case "SUBSTRING"
                '        sExpandedSymbol = ExpandSubstring(sSymbolAndFormat)

                '        '--- fixed items
                '    Case "TRIGGER_UNIQUE_ID"
                '        sExpandedSymbol = ExpandTriggerUniqueId(sSymbolAndFormat)

                '    Case "TRIGGER_TIMESTAMP"
                '        sExpandedSymbol = ExpandTriggerTimestamp(sSymbolAndFormat)


                'End Select

                '--- if symbol has been expanded then need to escape entity references
                If sExpandedSymbol <> LHS & sSymbolAndFormat & RHS Then
                    sExpandedSymbol = Escape(sExpandedSymbol)
                End If

                '--- update the hashtable
                _allSymbolsFound.Item(sSymbolAndFormat) = sExpandedSymbol

                '--- add to the ExpandedSymbols collection
                '_expandedSymbols.Add(sSymbol)
                _expandedSymbols.Add(sSymbolAndFormat)

            Else

                '--- add to the UnexpandedSymbols collection
                _unExpandedSymbols.Add(sSymbol)

            End If

        Next

    End Sub


    ''' <summary>
    ''' replace the symbols with the expanded symbols
    ''' </summary>
    ''' <remarks>SDCP 11-Oct-2006</remarks>
    Private Sub ReplaceWithExpansions()

        '--- get the unexpanded string
        _expandedString = String.Copy(_unExpandedString)

        '--- loop through the array of expanded symbols
        For Each expandedSymbol As String In _expandedSymbols

            '--- extract its symbol and expanded value
            Dim expandedValue As String = _allSymbolsFound.Item(expandedSymbol).ToString()

            '--- replace it in the input string
            _expandedString = _expandedString.Replace(LHS & expandedSymbol & RHS, expandedValue)

        Next

    End Sub


    ''' <summary>
    ''' Substring functionality, c.f. Mid$
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>Substring</returns>
    ''' <remarks>SDCP 10-Sep-2007</remarks>
    Private Function ExpandSubstring(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sStringToWorkOn As String = ""
        Dim lStartPos As Integer = 0
        Dim lNumChars As Integer = 0


        Try

            Select Case _scope

                Case Scope.PostTrigger

                    '--- get the parts
                    Try
                        sSymbolPart = Split(Symbol, "|")(0)
                        sStringToWorkOn = Split(Symbol, "|")(1)
                        lStartPos = CInt(Split(Symbol, "|")(2))
                        lNumChars = CInt(Split(Symbol, "|")(3))
                    Catch ex As Exception
                    End Try

                    '--- get value
                    sValue = Mid(sStringToWorkOn, lStartPos, lNumChars)

                Case Scope.PreTrigger
                    sValue = LHS & Symbol & RHS

            End Select

        Catch ex As Exception
        End Try

        '--- return the value
        Return sValue

    End Function



    ''' <summary>
    ''' Create a (formatted) date
    ''' #[DATEVALUE|isodatestring (yyyymmdd)|outputformat (dd-MMM-yyyy)|offsetdays (-1)|calendar (GBP)]#
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>Date as a string</returns>
    ''' <remarks>SDCP 12-Sep-2007</remarks>
    Private Function ExpandDateValue(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sISODateString As String = ""
        Dim sOutputFormat As String = "dd-MMM-yyyy"
        Dim lOffsetDays As Integer = 0
        Dim sCalendar As String = "GBP"

        Dim sHols As String = ""

        Dim dtInputDate As Date
        Dim dtOutputDate As Date


        Try

            Select Case _scope

                Case Scope.PreTrigger, Scope.PostTrigger

                    '--- get the parts
                    Try
                        sSymbolPart = Split(Symbol, "|")(0)
                        sISODateString = Split(Symbol, "|")(1)
                        sOutputFormat = Split(Symbol, "|")(2)
                        lOffsetDays = CInt(Split(Symbol, "|")(3))
                        sCalendar = Split(Symbol, "|")(4)
                    Catch ex As Exception
                    End Try

                    '--- get value:

                    '--- parse the ISO date input
                    dtInputDate = DateSerial(CInt(Mid(sISODateString, 1, 4)), CInt(Mid(sISODateString, 5, 2)), CInt(Mid(sISODateString, 7, 2)))

                    '--- try to read hols file
                    For Each de As DictionaryEntry In MyMessenger.Config.HolidayFiles
                        If de.Key.ToString = sCalendar Then
                            sHols = de.Value.ToString()
                            Exit For
                        End If
                    Next
                    Try
                        sHols = IO.File.OpenText(sHols).ReadToEnd
                    Catch ex As Exception
                        Dim sMessage As String = "#Unable to open GBP holiday file: " & sHols
                        MyMessenger.Log(Messenger.LogType.ErrorType, sMessage, "")
                    End Try

                    dtOutputDate = dtInputDate
                    Dim i As Integer = 0

                    'ML - 11dec07 - days fwd fix
                    'Adjust by offset days.  Move the date one day at a time so can check for hols or weekends at each move
                    If lOffsetDays < 0 Then

                        For i = -1 To lOffsetDays Step -1

                            dtOutputDate = dtOutputDate.AddDays(-1)
                            Do Until InStr(1, sHols, Format(dtOutputDate, "dd-MMM-yyyy"), CompareMethod.Text) = 0 _
                            And Weekday(dtOutputDate) <> 7 _
                            And Weekday(dtOutputDate) <> 1
                                dtOutputDate = dtOutputDate.AddDays(-1)
                            Loop

                        Next i

                    ElseIf lOffsetDays > 0 Then

                        For i = 1 To lOffsetDays

                            dtOutputDate = dtOutputDate.AddDays(1)
                            Do Until InStr(1, sHols, Format(dtOutputDate, "dd-MMM-yyyy"), CompareMethod.Text) = 0 _
                            And Weekday(dtOutputDate) <> 7 _
                            And Weekday(dtOutputDate) <> 1
                                dtOutputDate = dtOutputDate.AddDays(1)
                            Loop

                        Next i

                    End If


                    ''Adjust for weekend
                    'If Weekday(dtInputDate) = 2 Then
                    '    dtOutputDate = dtInputDate.AddDays(Math.Sign(lOffsetDays) * 2 + lOffsetDays)
                    'Else
                    '    dtOutputDate = dtInputDate.AddDays(lOffsetDays)
                    'End If
                    ''Keep moving the date backwards until find one that's not in the holiday file
                    'Do Until InStr(1, sHols, dtOutputDate.ToString("dd-MMM-yyyy"), CompareMethod.Text) = 0 And Weekday(dtOutputDate) <> 7 And Weekday(dtOutputDate) <> 1
                    '    dtOutputDate = dtOutputDate.AddDays(Math.Sign(lOffsetDays) * 1)
                    'Loop


                    '--- write the output
                    If sOutputFormat.ToLower = "panther" Then

                        sValue = ConvertToOBSDate(dtOutputDate)

                    Else

                        sValue = dtOutputDate.ToString(sOutputFormat)

                    End If


            End Select

        Catch ex As Exception
        End Try

        '--- return the value
        Return sValue

    End Function


    ''' <summary>
    ''' Create a (formatted) date
    ''' #[TIMEVALUE|isotimestring (HHmmss)|outputformat (HH-mm-ss)|offsetunit (hour)|offsetamount (-1)]#
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>Date as a string</returns>
    ''' <remarks>SDCP 12-Sep-2007</remarks>
    Private Function ExpandTimeValue(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sISOTimeString As String = ""
        Dim sOutputFormat As String = "HHmmss"
        Dim sOffsetUnit As String = "hour"
        Dim lOffsetAmount As Integer = 0

        Dim sHols As String = ""

        Dim dtInputTime As Date
        Dim dtOutputTime As Date


        Try

            Select Case _scope

                Case Scope.PreTrigger, Scope.PostTrigger

                    '--- get the parts
                    Try
                        sSymbolPart = Split(Symbol, "|")(0)
                        sISOTimeString = Split(Symbol, "|")(1)
                        sOutputFormat = Split(Symbol, "|")(2)
                        sOffsetUnit = Split(Symbol, "|")(3)
                        lOffsetAmount = CInt(Split(Symbol, "|")(4))
                    Catch ex As Exception
                    End Try

                    '--- get value:

                    '--- parse the ISO time input
                    dtInputTime = TimeSerial(CInt(Mid(sISOTimeString, 1, 2)), CInt(Mid(sISOTimeString, 3, 2)), CInt(Mid(sISOTimeString, 5, 2)))

                    '--- adjust for offset
                    dtOutputTime = GetAdjustedDate(dtInputTime, sOffsetUnit, lOffsetAmount)

                    '--- write the output
                    sValue = dtOutputTime.ToString(sOutputFormat)

            End Select

        Catch ex As Exception
        End Try

        '--- return the value
        Return sValue

    End Function


    ''' <summary>
    ''' rounds a time value *down* to the nearest 'n' units 
    ''' </summary>
    ''' <param name="Symbol">ROUNDTIME|151233|minute|5</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ExpandRoundTime(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sISOTimeString As String = ""
        Dim sRoundingUnit As String = "minute"
        Dim lRoundingAmount As Integer = 1

        Dim dtInputTime As Date
        Dim dtOutputTime As Date



        Try

            Select Case _scope

                Case Scope.PreTrigger, Scope.PostTrigger

                    '--- get the parts
                    Try
                        sSymbolPart = Split(Symbol, "|")(0)
                        sISOTimeString = Split(Symbol, "|")(1)
                        sRoundingUnit = Split(Symbol, "|")(2)
                        lRoundingAmount = CInt(Split(Symbol, "|")(3))
                    Catch ex As Exception
                    End Try

                    '--- get value:

                    '--- parse the ISO time input
                    dtInputTime = TimeSerial(CInt(Mid(sISOTimeString, 1, 2)), CInt(Mid(sISOTimeString, 3, 2)), CInt(Mid(sISOTimeString, 5, 2)))

                    '--- calculate the rounded time:

                    '--- find the timespan since 00:00:00
                    Dim t As TimeSpan = (dtInputTime.Subtract(DateTime.MinValue))

                    '--- calc how many rounding amounts will still be less than this timespan
                    Dim lCountOfRoundingAmounts As Integer = CInt(Int(t.TotalMinutes / lRoundingAmount))

                    '--- now just offset from 00:00:00 by this number of rounding amounts to get the rounded down time
                    dtOutputTime = DateTime.MinValue.Add(New TimeSpan(0, lCountOfRoundingAmounts * lRoundingAmount, 0))


                    '--- write the output
                    sValue = dtOutputTime.ToString("HHmmss")

            End Select

        Catch ex As Exception
        End Try





        '--- return the value
        Return sValue

    End Function


    Private Function ExpandToday(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sFormatPart As String = "dd-MMM-yyyy"

        Try

            '--- get the parts
            Try
                sSymbolPart = Split(Symbol, "|")(0)
                sFormatPart = Split(Symbol, "|")(1)
            Catch ex As Exception
            End Try

            '--- get Today
            sValue = FindDay(sSymbolPart.Replace("TRIGGER_", "")).ToString(sFormatPart)

        Catch ex As Exception
        End Try

        '--- return the value
        Return sValue

    End Function


    Private Function ExpandTriggerToday(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sFormatPart As String = "dd-MMM-yyyy"

        Try

            '--- get the parts
            Try
                sSymbolPart = Split(Symbol, "|")(0)
                sFormatPart = Split(Symbol, "|")(1)
            Catch ex As Exception
            End Try

            '--- get Today
            Select Case _scope

                Case Scope.PostTrigger
                    sValue = FindDay(sSymbolPart.Replace("TRIGGER_", "")).ToString(sFormatPart)

                Case Else
                    '

            End Select

        Catch ex As Exception
        End Try

        '--- return the value
        Return sValue

    End Function


    ''' <summary>
    ''' Expands NOW
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>Time as a string</returns>
    ''' <remarks></remarks>
    Private Function ExpandNow(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String
        Dim sFormatPart As String = "HH:mm:ss"


        Try

            '--- get the parts
            Try
                sSymbolPart = Split(Symbol, "|")(0)
                sFormatPart = Split(Symbol, "|")(1)
            Catch ex As Exception
            End Try

            '--- get Today
            sValue = Now().ToString(sFormatPart)

        Catch ex As Exception
        End Try

        Return sValue

    End Function

    ''' <summary>
    ''' Expands NOW
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>Time as a string</returns>
    ''' <remarks></remarks>
    Private Function ExpandTriggerNow(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String
        Dim sFormatPart As String = "HH:mm:ss"


        Try

            '--- get the parts
            Try
                sSymbolPart = Split(Symbol, "|")(0)
                sFormatPart = Split(Symbol, "|")(1)
            Catch ex As Exception
            End Try

            '--- get Now (post-trigger only)
            Select Case _scope

                Case Scope.PostTrigger
                    sValue = Now().ToString(sFormatPart)

                Case Else
                    '

            End Select

        Catch ex As Exception
        End Try

        Return sValue

    End Function


    ''' <summary>
    ''' Expands a fixed value for now
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>Time as a string</returns>
    ''' <remarks></remarks>
    Private Function ExpandFixedNow(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String
        Dim sFormatPart As String = "HH:mm:ss"


        Try

            '--- get the parts
            Try
                sSymbolPart = Split(Symbol, "|")(0)
                sFormatPart = Split(Symbol, "|")(1)
            Catch ex As Exception
            End Try

            '--- get Today
            sValue = _fixedNow.ToString(sFormatPart)

        Catch ex As Exception
        End Try

        Return sValue

    End Function



    ''' <summary>
    ''' Gets the trigger filename, for a FFW watcher
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>The filename and path</returns>
    ''' <remarks>returns just the symbol if filename not available</remarks>
    Private Function ExpandTriggerFilename(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS

        Try

            Select Case _scope
                Case Scope.PostTrigger
                    sValue = Path.GetFileName(_fileInfoTrigger.FullName)
                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function

    Private Function ExpandTriggerFilenameNoExtension(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS


        Try

            Select Case _scope
                Case Scope.PostTrigger
                    sValue = Path.GetFileNameWithoutExtension(_fileInfoTrigger.FullName)
                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function

    Private Function ExpandTriggerFilePath(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS

        Try

            Select Case _scope
                Case Scope.PostTrigger
                    sValue = Path.GetDirectoryName(_fileInfoTrigger.FullName)
                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function


    Private Function ExpandTriggerFileFullPath(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS

        Try

            Select Case _scope
                Case Scope.PostTrigger
                    sValue = _fileInfoTrigger.FullName
                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function


    Private Function ExpandTriggerFileLastModified(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sFormatPart As String = "yyyyMMdd_HHmmss"

        Try

            '--- get the parts
            Try
                sSymbolPart = Split(Symbol, "|")(0)
                sFormatPart = Split(Symbol, "|")(1)
            Catch ex As Exception
            End Try

            '--- get last modified date
            sValue = _fileInfoTrigger.LastWriteTime.ToString(sFormatPart)

        Catch ex As Exception
        End Try

        '--- return the value
        Return sValue

    End Function

    ''' <summary>
    ''' Expands the trigger's timestamp
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>Time as a string</returns>
    ''' <remarks></remarks>
    Private Function ExpandTriggerTimestamp(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String
        Dim sFormatPart As String = "HH:mm:ss"


        Try

            '--- get the parts        
            Select Case _scope
                Case Scope.PostTrigger
                    sSymbolPart = Split(Symbol, "|")(0)
                    sFormatPart = Split(Symbol, "|")(1)
                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

            '--- get Today
            sValue = _pulledTrigger.Timestamp.ToString(sFormatPart)

        Catch ex As Exception
        End Try

        Return sValue

    End Function


    ''' <summary>
    ''' Gets the trigger message Id for a QueueWatcher
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>The EMS Message ID</returns>
    ''' <remarks>returns just the symbol if filename not available</remarks>
    Private Function ExpandTriggerUniqueId(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS

        Try

            Select Case _scope
                Case Scope.PostTrigger
                    sValue = _pulledTrigger.UniqueID

                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function


    ''' <summary>
    ''' Gets the trigger message Id for a QueueWatcher
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>The EMS Message ID</returns>
    ''' <remarks>returns just the symbol if filename not available</remarks>
    Private Function ExpandTriggerMessageId(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS

        Try

            Select Case _scope
                Case Scope.PostTrigger
                    sValue = _gunMessageTrigger.MessageID

                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function


    ''' <summary>
    ''' Gets the trigger message Id for a QueueWatcher
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>The EMS Message ID</returns>
    ''' <remarks>returns just the symbol if filename not available</remarks>
    Private Function ExpandTriggerMessageNvp(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sNamePart As String = ""

        Try

            '--- get the parts
            Try
                sSymbolPart = Split(Symbol, "|")(0)
                sNamePart = Split(Symbol, "|")(1)
            Catch ex As Exception
            End Try

        Catch ex As Exception
        End Try

        Try

            Select Case _scope
                Case Scope.PostTrigger
                    '--- get the value
                    sValue = _gunMessageTrigger.GetNamedValue(sNamePart).ToString()

                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function

    ''' <summary>
    ''' Gets the trigger message Id for a QueueWatcher
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>The EMS Message ID</returns>
    ''' <remarks>returns just the symbol if filename not available</remarks>
    Private Function ExpandTriggerMessagePayload(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS

        Try

            Select Case _scope
                Case Scope.PostTrigger
                    '--- get the value
                    sValue = _gunMessageTrigger.Payload

                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function


    ''' <summary>
    ''' Gets job name
    ''' </summary>
    ''' <param name="Symbol">Symbol to be expanded</param>
    ''' <returns>expanded job name</returns>
    ''' <remarks>returns just the symbol if filename not available</remarks>
    Private Function ExpandJobName(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS

        Try

            Select Case _scope
                Case Scope.PostTrigger
                    sValue = _watcher.ActiveJob.JobXml.Attributes("name").Value

                Case Else
                    sValue = LHS & Symbol & RHS
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function


    ''' <summary>
    ''' Gets a queue alias
    ''' </summary>
    ''' <param name="Symbol">Symbol to be expanded</param>
    ''' <returns>expanded queue alias</returns>
    ''' <remarks>returns just the symbol if filename not available</remarks>
    Private Function ExpandQueueAlias(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sNamePart As String = ""

        Try

            sSymbolPart = Split(Symbol, "|")(0)
            sNamePart = Split(Symbol, "|")(1)

            Select Case sNamePart.ToLower
                Case "alltriggeredjobs"
                    sValue = MyMessenger.Config.AllTriggeredJobsQ
                Case "jobstodo"
                    sValue = MyMessenger.Config.JobsToDoQ
                Case "jobsinprogress"
                    sValue = MyMessenger.Config.JobsInProgressQ
                Case "jobsdone"
                    sValue = MyMessenger.Config.JobsDoneQ
                Case "status"
                    sValue = MyMessenger.Config.StatusQ
                Case "triggers"
                    sValue = MyMessenger.Config.TriggersQ
                Case "quarantine"
                    sValue = MyMessenger.Config.QuarantineQ
            End Select

        Catch ex As Exception

        End Try

        Return sValue

    End Function


    '''' <summary>
    '''' Gets job group name
    '''' </summary>
    '''' <param name="JobNode">The "job" node to be expanded</param>
    '''' <returns>expanded job node</returns>
    '''' <remarks>returns just the symbol if filename not available</remarks>
    'Private Function ExpandJobGroupName(ByVal JobNode As Xml.XmlNode) As String
    '    Dim sValue As String = LHS & "JOB_GROUP_NAME" & RHS

    '    Try

    '        Select Case _source

    '            Case "PreTriggering"
    '                sValue = JobNode.SelectSingleNode("ancestor::jobgroup").Attributes("name").Value

    '            Case Else
    '                '
    '                '

    '        End Select

    '    Catch ex As Exception

    '    End Try

    '    Return JobNode.InnerXml.Replace("#[JOB_GROUP_NAME]#", sValue)

    'End Function


    ''' <summary>
    ''' Gets a "unique" id
    ''' </summary>
    ''' <param name="Symbol">The full symbol text including formatting info</param>
    ''' <returns>a unique id</returns>
    ''' <remarks></remarks>
    Private Function ExpandUniqueId(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS

        Try

            sValue = CreateUniqueID()

        Catch ex As Exception

        End Try

        Return sValue

    End Function


    ''' <summary>
    ''' outputs the value of Environment variables
    ''' </summary>
    ''' <param name="Symbol"></param>
    ''' <returns></returns>
    ''' <remarks>SDCP 02-Oct-2007</remarks>
    Private Function ExpandEnviron(ByVal Symbol As String) As String
        Dim sValue As String = LHS & Symbol & RHS
        Dim sSymbolPart As String = ""
        Dim sFormatPart As String = ""

        Try

            '--- get the parts
            Try
                sSymbolPart = Split(Symbol, "|")(0)
                sFormatPart = Split(Symbol, "|")(1)
            Catch ex As Exception
            End Try

            '--- get Today
            sValue = Environ(sFormatPart)

        Catch ex As Exception
        End Try

        '--- return the value
        Return sValue

    End Function



    ''' <summary>
    '''Turn a recognised day expression into the actual date it represents.  Account for weekends and UK 
    ''' holidays.  
    ''' </summary>
    ''' <param name="DayExpression">"TODAY", "TODAY+1" or "TODAY-1"</param>
    ''' <returns>An adjusted date</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Private Function FindDay(ByVal DayExpression As String) As Date

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
                    FindDay = FindDay.AddDays(-1)
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
                    FindDay = FindDay.AddDays(1)
                Loop

            Case Else
                Err.Raise(-15031967, "FindDay", "Unable to determine date")

        End Select

    End Function





    '--- readonly properties
    Public ReadOnly Property ExpandedString() As String
        Get
            Return _expandedString
        End Get
    End Property
    Public ReadOnly Property ExpandedSymbols() As ArrayList
        Get
            Return _expandedSymbols
        End Get
    End Property
    Public ReadOnly Property UnexpandedSymbols() As ArrayList
        Get
            Return _unExpandedSymbols
        End Get
    End Property


End Class
