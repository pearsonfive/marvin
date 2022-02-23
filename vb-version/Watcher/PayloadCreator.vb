Imports System.IO
Imports OTIS


Public Class PayloadCreator


    Private _currentWatcher As WatcherBase
    Private _pulledTrigger As Object
    'Private _fileInfoTrigger As FileInfo
    'Private _gunMessageTrigger As GunMessage



    '--- CONSTRUCTORS
    Sub New(ByVal CurrentWatcher As WatcherBase, ByVal PulledTrigger As Object)

        ''--- cast to required type for watcher
        'Select Case True

        '    Case TypeOf CurrentWatcher Is FileFolderWatcher
        '        _fileInfoTrigger = CType(PulledTrigger, FileInfo)

        '    Case TypeOf CurrentWatcher Is FileFolderWatcher
        '        _gunMessageTrigger = CType(PulledTrigger, GunMessage)


        '    Case Else
        '        '
        '        '

        'End Select

        _currentWatcher = CurrentWatcher
        _pulledTrigger = PulledTrigger

    End Sub


    '**********************************************************
    '* create the payload and return it
    '* 
    '* 
    '*
    '* SDCP 28-Sep-2006
    '**********************************************************
    Public Function GetPayload() As String
        Dim sExpandedString As String

        '--- expand the symbolics
        sExpandedString = GetSymbolicallyExpandedString()

        '--- return the payload
        'Debug.Print(sExpandedString)
        Return sExpandedString

        'xTODO: Expose the unexpanded / expanded symbols collections?

    End Function


    '**********************************************************
    '* expand any #[SYMBOLICS|format-string]# in the triggerinfoxml
    '* 
    '* 
    '*
    '* SDCP 28-Sep-2006
    '**********************************************************
    Private Function GetSymbolicallyExpandedString() As String
        Dim PreExpansionXml As Xml.XmlNode
        Dim sPostExpansionString As String
        Dim oExpander As SymbolicsExpander


        '--- get the unexpanded Trigger xml
        PreExpansionXml = _currentWatcher.ActiveJob.JobXml

        '--- expand the symbolics
        oExpander = New SymbolicsExpander(PreExpansionXml, _currentWatcher, _pulledTrigger, SymbolicsExpander.Scope.PostTrigger)
        oExpander.Expand()
        sPostExpansionString = oExpander.ExpandedString

        '--- return the new string
        Return sPostExpansionString

    End Function

End Class
