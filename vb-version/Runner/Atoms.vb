Imports OTIS
Imports System.Threading

Public Class Atoms

#Region "Private"

    Private _Errs As Hashtable
    Private _Warnings As Hashtable

    Private _Atoms As Hashtable

    Private _Hash As String
    Private _Timeout As Long
    Private _xPayload As Xml.XmlDocument

    Private _GoodFinish As Integer
    Private _BadFinish As Integer

    ''' <summary>
    ''' Start the atom on a new thread
    ''' </summary>
    ''' <param name="_Atom"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function StartAtom(ByVal _Atom As Atom) As Boolean

        Dim t As Thread
        t = New Thread(AddressOf AtomRunner)
        t.Start(_Atom)

    End Function

    ''' <summary>
    ''' Parameters is a single entry hashtable with key as the exe path, value as switches string.  
    ''' When atom finishes, either increment the good or bad finish, depending on whether atom
    ''' finished normally or was terminated
    ''' </summary>
    ''' <param name="_Atom"></param>
    ''' <remarks></remarks>
    Private Sub AtomRunner(ByVal _Atom As Object)

        'Read atom settings.  Always pass hash to runner, may as well do it here
        Dim myAtom As Atom = CType(_Atom, Atom)
        Dim CommandLine As String = myAtom.CommandLine
        Dim Switches As String = myAtom.Switches
        Switches += " /hash=" & DQt(_Hash)
        Switches += " /environment=" & DQt(MySwitches("environment"))

        Dim proc As Process = New Process
        With proc

            .StartInfo.FileName = DQt(CommandLine)
            .StartInfo.Arguments = Switches
            .StartInfo.UseShellExecute = True
            'MyMessenger.Log(Messenger.LogType.StatusType, DQt(CommandLine) & " " & Switches)
            .Start()
            Dim fHasRunProperly As Boolean = .WaitForExit(_Timeout * 1000)
            If fHasRunProperly Then
                _GoodFinish += 1
            Else
                Try
                    .Kill()
                Catch
                End Try
                _BadFinish += 1
                MyMessenger.AtomInfo += "<status>Runner aborted the atom</status>"
                MyMessenger.AtomInfo += "<commandline>" & CommandLine & "</commandline>"
                MyMessenger.AtomInfo += "<switches>" & Switches & "</switches>"
                MyMessenger.AtomInfo += "<timeout>" & _Timeout * 1000.ToString & "</timeout>"
            End If

        End With

    End Sub

#End Region

#Region "Public"

    Public ReadOnly Property Errors() As Hashtable
        Get
            Return _Errs
        End Get
    End Property

    Public ReadOnly Property Warnings() As Hashtable
        Get
            Return _Warnings
        End Get
    End Property

    Public ReadOnly Property Atoms() As Hashtable
        Get
            Return _Atoms
        End Get
    End Property

    Public ReadOnly Property Timeout() As Long
        Get
            Return _Timeout
        End Get
    End Property

    ''' <summary>
    ''' From the message payload work out which atoms need to be run. Call SetAtomDefaults to set 
    ''' each atom's default settings; call SetAtomOverrides to set job-specific settings.
    ''' </summary>
    ''' <param name="TriggerMessage"></param>
    ''' <param name="ConfigDoc"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal TriggerMessage As GunMessage, ByVal ConfigDoc As Xml.XmlDocument)

        _Errs = New Hashtable
        _Warnings = New Hashtable
        _Atoms = New Hashtable

        'Get the hash from the gun message
        _Hash = TriggerMessage.Hash

        'Load the payload xml
        _xPayload = New Xml.XmlDataDocument
        _xPayload.LoadXml(TriggerMessage.Payload)

        'get the name of each atom to run; setup an atom object containing all settings for the named
        'atom; write each object into a hashtable.
        Dim xNodeList As Xml.XmlNodeList
        xNodeList = _xPayload.DocumentElement.SelectNodes("//atom")
        For Each xNode As Xml.XmlNode In xNodeList

            'Null atom job - do the WhenFinished piece only
            If xNode.Attributes("name").Value.ToLower = "null" And xNodeList.Count = 1 Then
                _Warnings.Add("Null", "Null atom job")


                '--- drop a crumb to show this
                MyCrumbler = New Crumbler(TriggerMessage.Pairs)
                MyCrumbler.DropCrumb("NullAtom", GetBaseRunnerInfo(TriggerMessage))
            
                DoWhenFinishedActions(TriggerMessage.Payload, MyMessenger)
                Exit Sub
            End If

            Dim myAtom As Atom = New Atom(ConfigDoc, xNode)
            If myAtom.Errors.Count = 0 Then
                _Atoms.Add(_Atoms.Count, myAtom)
                _Timeout = System.Math.Max(myAtom.Timeout, _Timeout)

                'Log any warnings that were generated - atom was still added
                For Each de As DictionaryEntry In myAtom.Warnings
                    Dim sLogMessage As String = de.Value.ToString
                    MyMessenger.Log(Messenger.LogType.StatusType, sLogMessage, _Hash)
                Next

            Else
                'Log any errors that were generated - atom was not added
                For Each de As DictionaryEntry In myAtom.Errors
                    Dim sLogMessage As String = de.Value.ToString
                    MyMessenger.Log(Messenger.LogType.ErrorType, sLogMessage, _Hash)
                Next

            End If

        Next

    End Sub

    ''' <summary>
    ''' Using the modified settings and switches, run each atom on a new thread.  Return true
    ''' only if ALL atoms finished properly
    ''' </summary>
    ''' <returns>True if atom completed adequately; false if timed out</returns>
    ''' <remarks>Martin Leyland, December 2006</remarks>
    Public Function Run() As Boolean

        'Possible for there to be no atoms here, in which case just move on
        For Each de As DictionaryEntry In _Atoms

            'Start the atom.  Each atom will run on a new thread
            Dim myAtom As Atom = CType(de.Value, Atom)
            StartAtom(myAtom)

        Next

        'each thread will either increment good or bad finish counter, according to whether atom finished
        'normally or was timed out.  Keep checking until the "all atoms" timeout (=max atom timeout)
        Dim dStopChecking As Date = Now.AddSeconds(_Timeout)
        Do Until _GoodFinish + _BadFinish = _Atoms.Count Or Now > dStopChecking
            Thread.Sleep(500)
        Loop

        'Once all atoms have finished, return True if all finished OK, otherwise false
        If _BadFinish > 0 Or _GoodFinish + _BadFinish <> _Atoms.Count Then
            Return False
        Else
            Return True
        End If

    End Function

#End Region

End Class
