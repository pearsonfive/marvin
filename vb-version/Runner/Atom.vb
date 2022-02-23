Public Class Atom

#Region "Private"

    Private _CommandLine As String
    Private _Switches As String
    Private _SwitchesHT As Hashtable
    Private _Timeout As Long

    Private _Errs As Hashtable
    Private _Warnings As Hashtable

    ''' <summary>
    ''' Read from the app config the default settings and switches for the atom
    ''' </summary>
    ''' <param name="AtomName"></param>
    ''' <param name="ApplicationConfig"></param>
    ''' <remarks></remarks>
    Private Sub SetAtomDefaults(ByVal AtomName As String, ByVal ApplicationConfig As Xml.XmlDocument)

        Try

            'Set atom timeout.
            Dim xNode As Xml.XmlNode, xPath As String
            xPath = "//atomapps/atom [@name=" & Qt(AtomName) & "]/timeout"
            xNode = ApplicationConfig.SelectSingleNode(xPath)
            Try
                _Timeout = CLng(xNode.InnerText)
            Catch ex As Exception
                _Warnings.Add(_Errs.Count + 1, "#Atom timeout not read properly from xml config")
                _Timeout = 0
            End Try

            'Get base cmd line for the atom
            xPath = "//atomapps/atom[@name=" & Qt(AtomName) & "]/commandline/base"
            xNode = ApplicationConfig.SelectSingleNode(xPath)
            _CommandLine = xNode.InnerText
            Try
                If xNode.Attributes("type").Value = "relative" Then
                    _CommandLine = System.IO.Path.Combine(MyMessenger.Config.EnvironmentAppPath, _CommandLine)
                End If
            Catch
                'there is no "type" attribute, so must be absolute - leave cmd line intact
            End Try

            'Find all the switches
            Dim xNodeList As Xml.XmlNodeList, sLHS As String = "", sRHS As String = ""
            _SwitchesHT = New Hashtable
            xPath = "//atomapps/atom [@name=" & Qt(AtomName) & "]/commandline/switches/switch"
            xNodeList = ApplicationConfig.SelectNodes(xPath)
            For Each xNode In xNodeList
                sLHS = " /" & xNode.SelectSingleNode("lhs").InnerText
                sRHS = "=" & DQt(xNode.SelectSingleNode("rhs").InnerText)
                _SwitchesHT.Add(sLHS, sRHS)
            Next

        Catch ex As Exception
            Dim sMessage As String = AtomName & " not found in AppConfig.xml"
            MyMessenger.Log(Messenger.LogType.ErrorType, sMessage)
            End

        End Try

    End Sub

    ''' <summary>
    ''' Read from the payload any settings that should be overridden for this particular job
    ''' </summary>
    ''' <param name="AtomNode"></param>
    ''' <remarks></remarks>
    Private Sub SetAtomOverRides(ByVal AtomNode As Xml.XmlNode)

        Dim xpath As String = "descendant::*"
        Dim xNodeList As Xml.XmlNodeList = AtomNode.SelectNodes(xpath)

        'examine all descendents of the node, get switches, timeout and cmd line
        For Each xNode As Xml.XmlNode In xNodeList
            Dim sLHS As String = "", sRHS As String = ""
            If xNode.Name = "switch" Then
                sLHS = " /" & xNode.SelectSingleNode("lhs").InnerText
                sRHS = "=" & DQt(xNode.SelectSingleNode("rhs").InnerText)
                Try
                    _SwitchesHT.Remove(sLHS)
                Catch
                End Try
                _SwitchesHT.Add(sLHS, sRHS)
            ElseIf xNode.Name = "base" Then
                _CommandLine = xNode.InnerText
            ElseIf xNode.Name = "timeout" Then
                Try
                    _Timeout = CLng(xNode.InnerText)
                Catch
                End Try
            End If
        Next

    End Sub

    ''' <summary>
    ''' Concatenate a hashtable, with separators for name/values and each entry
    ''' </summary>
    ''' <param name="_HashTable"></param>
    ''' <returns>concat string</returns>
    ''' <remarks></remarks>
    Private Function HashtableToString(ByVal _HashTable As Hashtable) As String

        Dim sResult As String = ""
        For Each de As DictionaryEntry In _HashTable
            Try
                sResult += de.Key.ToString() & de.Value.ToString
            Catch
            End Try
        Next
        Return sResult

    End Function

#End Region

#Region "Public"

    Public ReadOnly Property Timeout() As Long
        Get
            Return _Timeout
        End Get
    End Property

    Public ReadOnly Property CommandLine() As String
        Get
            Return _CommandLine
        End Get
    End Property

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

    Public ReadOnly Property Switches() As String
        Get
            Return _Switches
        End Get
    End Property

    ''' <summary>
    ''' Set default atom timeout, cmd line and switches; override these if specified in the payload
    ''' </summary>
    ''' <param name="ApplicationConfig"></param>
    ''' <param name="AtomNode"></param>
    ''' <remarks></remarks>
    Sub New(ByVal ApplicationConfig As Xml.XmlDocument, ByVal AtomNode As Xml.XmlNode)

        _Errs = New Hashtable
        _Warnings = New Hashtable

        'Set defaults for this atom - timeout, cmd line & switches
        SetAtomDefaults(AtomNode.Attributes("name").Value, ApplicationConfig)

        'Some of these will be overridden in the job definition (the payload).
        SetAtomOverRides(AtomNode)

        'Transform switches hashtable into a string
        _Switches = HashtableToString(_SwitchesHT)

    End Sub

#End Region

End Class
