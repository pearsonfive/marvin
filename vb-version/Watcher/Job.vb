Public Class Job

    Private _jobXml As Xml.XmlNode
    Private _hash As String
    Private _watcherType As String
    Private _name As String
    Private _triggered As Boolean
    '
    Private _trigger As Trigger
    Private _combinationLogic As String



    Public Property JobXml() As Xml.XmlNode
        Get
            Return _jobXml
        End Get
        Friend Set(ByVal value As Xml.XmlNode)
            _jobXml = value
        End Set
    End Property


    Public ReadOnly Property Trigger() As Trigger
        Get
            Return _trigger
        End Get
    End Property


    Public Sub BuildTrigger()
        '-- create the trigger object
        _trigger = CreateTrigger()
    End Sub


    ''' <summary>
    ''' creates the trigger object for this job
    ''' </summary>
    ''' <returns>A Trigger object</returns>
    ''' <remarks>SDCP 26-Oct-2006</remarks>
    Private Function CreateTrigger() As Trigger
        Dim tgr As New Trigger
        Dim nTrigger As Xml.XmlNode

        nTrigger = _jobXml.SelectSingleNode("trigger")

        '--- set it up
        tgr.TriggerInfoXml = nTrigger
        tgr.WatcherType = nTrigger.SelectSingleNode("watcherbase/@type").Value

        '--- return it
        Return tgr

    End Function


    Public Property Hash() As String
        Get
            Return _hash
        End Get
        Friend Set(ByVal value As String)
            _hash = value
        End Set
    End Property


    Public Property WatcherType() As String
        Get
            Return _watcherType
        End Get
        Friend Set(ByVal value As String)
            _watcherType = value
        End Set
    End Property


    Public Property Name() As String
        Get
            Return _name
        End Get
        Friend Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Triggered() As Boolean
        Get
            Return _triggered
        End Get
        Set(ByVal value As Boolean)
            _triggered = value
        End Set
    End Property


    ''' <summary>
    ''' Replace instances of local variables in the JobDef
    ''' </summary>
    ''' <remarks>SDCP 16-Apr-2007</remarks>
    Public Sub ExtractLocalVariables()
        Dim sVariableName As String
        Dim sVariableInitialValue As String
        Dim sSymbolToReplace As String

        Dim sOriginalXML As String

        '--- get the current Job Xml
        sOriginalXML = _jobXml.InnerXml

        For Each ndLocal As Xml.XmlNode In _jobXml.SelectNodes("locals/variable")
            Try

                sVariableName = ndLocal.Attributes("name").Value
                sVariableInitialValue = ndLocal.InnerXml

                '--- create the symbol to look for and replace
                sSymbolToReplace = "$" & sVariableName & "$"

                '--- replace it in the input string
                sOriginalXML = sOriginalXML.Replace(sSymbolToReplace, sVariableInitialValue)

            Catch ex As Exception

                MyMessenger.AtomInfo += "<error>Local Variable incorrectly defined</error>"
                MyMessenger.AtomInfo += "<localvariablexml>" & ndLocal.OuterXml & "</localvariablexml>"
                Err.Raise(vbObjectError + 1001)

            End Try

        Next

        '--- reload the XML
        _jobXml.InnerXml = sOriginalXML

    End Sub

End Class
