Public Class Config

    Public Structure QueueName
        Public PrettyName As String
        Public ActualName As String
        Public DisplayOrder As Integer
    End Structure


    'Tibco settings
    Private _EmsServer As String
    Private _EmsUsername As String
    Private _EmsPassword As String

    'Queue names
    Private _AllTriggeredJobs As String
    Private _JobsToDo As String
    Private _JobsInProgress As String
    Private _JobsDone As String
    Private _Status As String
    Private _Triggers As String
    Private _Quarantine As String
    Private _Trail As String

    'Private _AllQueues As Hashtable
    Private _AllQueues As SortedList(Of Integer, Config.QueueName)

    '--- Topic names
    Private _Clock As String

    'DB settings
    Private _DBName As String
    Private _DBServer As String

    'App config
    Private _RPMConfigFile As String

    'Poll frequency
    Private _PollFrequency As Integer

    '--- AppConfig.xml DomDocument
    Private _DomDoc As Xml.XmlDocument

    '--- available symbols
    Private _AvailableSymbols As Hashtable
    Private _ContextualSymbols As Hashtable
    Private _NonContextualSymbols As Hashtable

    'Holiday files
    Private _HolidayFiles As Hashtable

    'Error
    Private _Errors As Hashtable

    'Admin users
    Private _AdminUsers As Hashtable

    '--- age of messages to consume in Runner
    Private _RunnerMessageAge As Integer
    Private _MaxAuditPoints As Integer

    '--- app path for current environment
    Private _EnvironmentAppPath As String

    '--- service account
    Private _ServiceAccountUsername As String
    Private _ServiceAccountPassword As String


    ''' <summary>
    ''' Initialising the class exposes properties for all of the application level settings
    ''' </summary>
    ''' <param name="ConfigFile"></param>
    ''' <param name="CallingApp"></param>
    ''' <remarks>Martin Leyland, November 2006</remarks>
    ''' 
    Public Sub New(ByVal ConfigFile As String, ByVal CallingApp As String)

        Dim xNode As Xml.XmlNode
        Dim xmlPreBootstrapping As Xml.XmlDocument

        'Set up an errors collection
        Dim Errs As New Hashtable

        'Allow the passed config to be retrieved
        _RPMConfigFile = ConfigFile

        '--- load the app config file and bootstrap some includes
        xmlPreBootstrapping = New Xml.XmlDocument
        xmlPreBootstrapping.Load(ConfigFile)

        '--- check environment switch against available environments
        Dim xPath As String = "//RPMConfig/apppaths/env[@name='" & MySwitches("environment") & "']"
        xNode = xmlPreBootstrapping.DocumentElement.SelectSingleNode(xPath)
        If xNode Is Nothing Then
            Dim sMessage As String = "No settings found for this environment - " & MySwitches("environment")
            LastDitchLog(sMessage)
        Else
            _EnvironmentAppPath = xNode.Attributes("path").Value
        End If


        '--- bootstrapping the includes
        _DomDoc = New Xml.XmlDataDocument
        _DomDoc.LoadXml(BootstrapIncludes(xmlPreBootstrapping.OuterXml))

        'Set up some properties

        Try

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/server")
            If xNode Is Nothing Then
                Errs.Add(1, "Config error: //RPMConfig/environment/server not found")
            Else
                _EmsServer = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/username")
            If xNode Is Nothing Then
                Errs.Add(2, "Config error: //RPMConfig/environment/username not found")
            Else
                _EmsUsername = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/password")
            If xNode Is Nothing Then
                Errs.Add(3, "Config error: //RPMConfig/environment/password not found")
            Else
                _EmsPassword = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/queues/queue[@name='AllTriggeredJobs']")
            If xNode Is Nothing Then
                Errs.Add(4, "Config error: //RPMConfig/environment/queues/queue[@name='AllTriggeredJobs'] not found")
            Else
                _AllTriggeredJobs = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/queues/queue[@name='JobsToDo']")
            If xNode Is Nothing Then
                Errs.Add(5, "Config error: //RPMConfig/environment/queues/queue[@name='JobsToDo'] not found")
            Else
                _JobsToDo = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/queues/queue[@name='JobsInProgress']")
            If xNode Is Nothing Then
                Errs.Add(6, "Config error: //RPMConfig/environment/queues/queue[@name='JobsInProgress'] not found")
            Else
                _JobsInProgress = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/queues/queue[@name='JobsDone']")
            If xNode Is Nothing Then
                Errs.Add(7, "Config error: //RPMConfig/environment/queues/queue[@name='JobsDone'] not found")
            Else
                _JobsDone = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/queues/queue[@name='Status']")
            If xNode Is Nothing Then
                Errs.Add(8, "Config error: //RPMConfig/environment/queues/queue[@name='Status'] not found")
            Else
                _Status = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/queues/queue[@name='Triggers']")
            If xNode Is Nothing Then
                Errs.Add(9, "Config error: //RPMConfig/environment/queues/queue[@name='Triggers'] not found")
            Else
                _Triggers = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/queues/queue[@name='Quarantine']")
            If xNode Is Nothing Then
                Errs.Add(9, "Config error: //RPMConfig/environment/queues/queue[@name='Quarantine'] not found")
            Else
                _Quarantine = xNode.InnerText
            End If


            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/queues/queue[@name='Trail']")
            If xNode Is Nothing Then
                Errs.Add(9, "Config error: //RPMConfig/environment/queues/queue[@name='Trail'] not found")
            Else
                _Trail = xNode.InnerText
            End If


            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/topics/topic[@name='Clock']")
            If xNode Is Nothing Then
                Errs.Add(8, "Config error: //RPMConfig/environment/topics/topic[@name='Clock'] not found")
            Else
                _Clock = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/database/name")
            If xNode Is Nothing Then
                Errs.Add(10, "Config error: //RPMConfig/environment/database/name not found")
            Else
                _DBName = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/database/server")
            If xNode Is Nothing Then
                Errs.Add(11, "Config error: //RPMConfig/environment/database/server not found")
            Else
                _DBServer = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/coreapps/application[@name='" & CallingApp & "']/pollinterval")
            If xNode Is Nothing Then
                _PollFrequency = 0
            Else
                _PollFrequency = CInt(xNode.InnerText)
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/coreapps/application[@name='" & CallingApp & "']/messageage")
            If xNode Is Nothing Then
                _RunnerMessageAge = 0
            Else
                _RunnerMessageAge = CInt(xNode.InnerText)
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/environment/coreapps/application[@name='runner']/maxauditpoints")
            If xNode Is Nothing Then
                _MaxAuditPoints = 0
            Else
                _MaxAuditPoints = CInt(xNode.InnerText)
            End If

            '--- Available Symbols
            _AvailableSymbols = New Hashtable
            For Each xNode In _DomDoc.SelectNodes("//environment/symbols/symbol")
                _AvailableSymbols.Add(xNode.InnerText, xNode.InnerText)
            Next

            '--- Contextual Symbols
            _ContextualSymbols = New Hashtable
            For Each xNode In _DomDoc.SelectNodes("//environment/symbols/symbol[@type='contextual']")
                _ContextualSymbols.Add(xNode.InnerText, xNode.InnerText)
            Next

            '--- Non-Contextual Symbols
            _NonContextualSymbols = New Hashtable
            For Each xNode In _DomDoc.SelectNodes("//environment/symbols/symbol[@type='simple']")
                _NonContextualSymbols.Add(xNode.InnerText, xNode.InnerText)
            Next

            'Holiday files
            _HolidayFiles = New Hashtable
            For Each xNode In _DomDoc.SelectNodes("//environment/holidayfiles/file")
                _HolidayFiles.Add(xNode.Attributes("ccy").Value, xNode.InnerText)
            Next

            'List of admin users
            _AdminUsers = New Hashtable
            For Each xNode In _DomDoc.SelectNodes("//environment/adminusers/user")
                _AdminUsers.Add(_AdminUsers.Count + 1, xNode.InnerText)
            Next

            '--- sorted list of queues
            _AllQueues = New SortedList(Of Integer, Config.QueueName)
            For Each xNode In _DomDoc.SelectNodes("//RPMConfig/environment/queues/queue")
                Dim qn As New Config.QueueName
                With qn
                    .PrettyName = xNode.Attributes("name").Value
                    .ActualName = xNode.InnerXml
                    .DisplayOrder = CInt(xNode.Attributes("displayorder").Value)
                End With
                _AllQueues.Add(qn.DisplayOrder, qn)
            Next

            '--- service account
            xNode = _DomDoc.SelectSingleNode("//RPMConfig/serviceaccount/username")
            If xNode Is Nothing Then
                Errs.Add(8, "Config error: //RPMConfig/serviceaccount/username not found")
            Else
                _ServiceAccountUsername = xNode.InnerText
            End If

            xNode = _DomDoc.SelectSingleNode("//RPMConfig/serviceaccount/password")
            If xNode Is Nothing Then
                Errs.Add(8, "Config error: //RPMConfig/serviceaccount/password not found")
            Else
                _ServiceAccountPassword = xNode.InnerText
            End If



            '--- pop the errors variable
            _Errors = Errs

        Catch ex As Exception

        End Try

    End Sub

    Public ReadOnly Property EmsServer() As String
        Get
            Return _EmsServer
        End Get
    End Property

    Public ReadOnly Property EmsUsername() As String
        Get
            Return _EmsUsername
        End Get
    End Property

    Public ReadOnly Property EmsPassword() As String
        Get
            Return _EmsPassword
        End Get
    End Property

    Public ReadOnly Property AllTriggeredJobsQ() As String
        Get
            Return _AllTriggeredJobs
        End Get
    End Property

    Public ReadOnly Property JobsToDoQ() As String
        Get
            Return _JobsToDo
        End Get
    End Property

    Public ReadOnly Property JobsInProgressQ() As String
        Get
            Return _JobsInProgress
        End Get
    End Property

    Public ReadOnly Property JobsDoneQ() As String
        Get
            Return _JobsDone
        End Get
    End Property

    Public ReadOnly Property StatusQ() As String
        Get
            Return _Status
        End Get
    End Property

    Public ReadOnly Property TriggersQ() As String
        Get
            Return _Triggers
        End Get
    End Property

    Public ReadOnly Property QuarantineQ() As String
        Get
            Return _Quarantine
        End Get
    End Property

    Public ReadOnly Property BreadcrumbsQ() As String
        Get
            Return _Trail
        End Get
    End Property

    Public ReadOnly Property ClockTopic() As String
        Get
            Return _Clock
        End Get
    End Property

    Public ReadOnly Property DBName() As String
        Get
            Return _DBName
        End Get
    End Property

    Public ReadOnly Property DBServer() As String
        Get
            Return _DBServer
        End Get
    End Property

    Public ReadOnly Property PollFrequency() As Integer
        Get
            Return _PollFrequency
        End Get
    End Property

    Public ReadOnly Property RunnerMessageAge() As Integer
        Get
            Return _RunnerMessageAge
        End Get
    End Property

    Public ReadOnly Property MaxAuditPoints() As Integer
        Get
            Return _MaxAuditPoints
        End Get
    End Property

    Public ReadOnly Property Errors() As Hashtable
        Get
            Return _Errors
        End Get
    End Property

    Public ReadOnly Property RPMConfigFile() As String
        Get
            Return _RPMConfigFile
        End Get
    End Property

    Public ReadOnly Property AvailableSymbols() As Hashtable
        Get
            Return _AvailableSymbols
        End Get
    End Property

    Public ReadOnly Property ContextualSymbols() As Hashtable
        Get
            Return _ContextualSymbols
        End Get
    End Property

    Public ReadOnly Property NonContextualSymbols() As Hashtable
        Get
            Return _NonContextualSymbols
        End Get
    End Property

    Public ReadOnly Property HolidayFiles() As Hashtable
        Get
            Return _HolidayFiles
        End Get
    End Property

    Public ReadOnly Property AdminUsers() As Hashtable
        Get
            Return _AdminUsers
        End Get
    End Property

    Public ReadOnly Property DomDoc() As Xml.XmlDocument
        Get
            Return _DomDoc
        End Get
    End Property

    'Public ReadOnly Property AllQueues() As Hashtable
    '    Get
    '        Return _AllQueues
    '    End Get
    'End Property

    Public ReadOnly Property AllQueues() As SortedList(Of Integer, Config.QueueName)
        Get
            Return _AllQueues
        End Get
    End Property

    Public ReadOnly Property EnvironmentAppPath() As String
        Get
            Return _EnvironmentAppPath
        End Get
    End Property

    Public ReadOnly Property ServiceAccountUsername() As String
        Get
            Return _ServiceAccountUsername
        End Get
    End Property

    Public ReadOnly Property ServiceAccountPassword() As String
        Get
            Return _ServiceAccountPassword
        End Get
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

