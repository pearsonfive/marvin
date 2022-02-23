Imports OTIS
Imports System.Xml


Public Class Crumbler

    Private _TrailLabel As String = ""
    Private _TrailDescription As String = ""
    Private _UID As String = ""
    Private _EventName As String = ""
    Private _Info As XmlNode



    ''' <summary>
    ''' Creates a blank Crumbler ready for top-level manipulation
    ''' <param name="Label"></param>
    ''' </summary>
    ''' <remarks>SDCP 22-Jan-2008</remarks>
    Public Sub New(ByVal TrailLabel As String, ByVal TrailDescription As String)

        '--- store the info
        _TrailLabel = TrailLabel
        _TrailDescription = TrailDescription

        '--- create the UID
        _UID = GetNewUID()

        '--- reset the current event name (just in case)
        _EventName = ""

    End Sub


    ''' <summary>
    ''' Sets up the Crumbler from the payload of the job
    ''' </summary>
    ''' <param name="PayloadXml"></param>
    ''' <remarks>SDCP 25-Jan-2008</remarks>
    Public Sub New(ByVal PayloadXml As XmlNode)

        '--- see if a <trail> node exists
        Dim xTrailNode As XmlNode = Nothing
        Try
            xTrailNode = PayloadXml.SelectSingleNode("trail")
        Catch ex As Exception
        End Try


        If xTrailNode IsNot Nothing Then

            '--- need to extract the Label and description
            _TrailLabel = xTrailNode.SelectSingleNode("label").InnerText
            _TrailDescription = xTrailNode.SelectSingleNode("description").InnerText

            '--- and create the UID
            _UID = GetNewUID()

        End If

    End Sub


    ''' <summary>
    ''' Sets up the Crumbler from ?ADS info
    ''' </summary>
    ''' <remarks>SDCP 22-Jan-2008</remarks>
    Public Sub New(ByVal TrailLabel As String, ByVal TrailDescription As String, ByVal TrailUID As String)

        '--- store the info
        _TrailLabel = TrailLabel
        _TrailDescription = TrailDescription
        _UID = TrailUID

        '--- reset the current event name (just in case)
        _EventName = ""

    End Sub


    ''' <summary>
    ''' sets up the Crumbler from NVPs
    ''' </summary>
    ''' <param name="NVPs"></param>
    ''' <remarks>SDCP 01-Feb-2008</remarks>
    Public Sub New(ByVal NVPs As Hashtable)

        '--- try to load from NVP hashtable
        If NVPs.ContainsKey("BC_TrailLabel") Then _TrailLabel = NVPs("BC_TrailLabel").ToString
        If NVPs.ContainsKey("BC_TrailDescription") Then _TrailDescription = NVPs("BC_TrailDescription").ToString
        If NVPs.ContainsKey("BC_UID") Then _UID = NVPs("BC_UID").ToString

    End Sub


    ''' <summary>
    ''' The Label for the job chain.
    ''' Will be persisted to all crumbs for this Crumbler.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 22-Jan-2008</remarks>
    Public Property TrailLabel() As String
        Get
            Return _TrailLabel
        End Get
        Set(ByVal value As String)
            _TrailLabel = value
        End Set
    End Property


    ''' <summary>
    ''' The Description for the job chain.
    ''' Will be persisted to all crumbs for this Crumbler.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 22-Jan-2008</remarks>
    Public Property TrailDescription() As String
        Get
            Return _TrailDescription
        End Get
        Set(ByVal value As String)
            _TrailDescription = value
        End Set
    End Property


    ''' <summary>
    ''' Unique identifier for this job (from which no meaning can be inferred)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 22-Jan-2008</remarks>
    Public Property UID() As String
        Get
            Return _UID
        End Get
        Set(ByVal value As String)
            _UID = value
        End Set
    End Property


    ''' <summary>
    ''' The current breadcrumb trail
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 22-Jan-2008</remarks>
    Public Property EventName() As String
        Get
            Return _EventName
        End Get
        Set(ByVal value As String)
            _EventName = value
        End Set
    End Property


    ''' <summary>
    ''' The Info xml document for the current crumb
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>SDCP 25-Jan-2008</remarks>
    Public Property Info() As XmlNode
        Get
            Return _Info
        End Get
        Set(ByVal value As XmlNode)
            _Info = value
        End Set
    End Property


    ''' <summary>
    ''' Writes the crumb to the Breadcrumbs queue
    ''' </summary>
    ''' <param name="EventName"></param>
    ''' <param name="Info"></param>
    ''' <remarks>SDCP 22-Jan-2008</remarks>
    Public Sub DropCrumb(ByVal EventName As String, ByVal Info As XmlNode)

        '--- only continue if we have a valid trail, e.g. we have a TrailLabel and UID and Description
        If (_TrailLabel <> "" And _TrailDescription <> "" And _UID <> "") Then

            '--- update the current trail
            _EventName = EventName

            Dim nvp As Hashtable = New Hashtable

            '--- write breadcrumb nvps
            With nvp
                .Add("BC_TrailLabel", _TrailLabel)
                .Add("BC_TrailDescription", _TrailDescription)
                .Add("BC_UID", _UID)
                .Add("BC_EventName", _EventName)
                .Add("BC_Info", Info.OuterXml)
            End With

            '--- post the message
            MyMessenger.PostMessage(MyMessenger.Config.BreadcrumbsQ, "", "", "", nvp)

        End If

    End Sub


    ''' <summary>
    ''' Writes the crumb to the Breadcrumbs queue
    ''' </summary>
    ''' <param name="EventName"></param>
    ''' <param name="InfoText"></param>
    ''' <remarks>SDCP 18-Feb-2008</remarks>
    Public Sub DropCrumb(ByVal EventName As String, ByVal InfoText As String)
        Dim Info As XmlDocument = Nothing


        '--- only continue if we have a valid trail, e.g. we have a TrailLabel and UID and Description
        If (_TrailLabel <> "" And _TrailDescription <> "" And _UID <> "") Then

            '--- create the <info> node from the InfoText
            Info = New XmlDocument()
            Info.LoadXml("<info/>")
            Info.DocumentElement.InnerXml = InfoText

            '--- update the current trail
            _EventName = EventName

            Dim nvp As Hashtable = New Hashtable

            '--- write breadcrumb nvps
            With nvp
                .Add("BC_TrailLabel", _TrailLabel)
                .Add("BC_TrailDescription", _TrailDescription)
                .Add("BC_UID", _UID)
                .Add("BC_EventName", _EventName)
                .Add("BC_Info", Info.OuterXml)
            End With

            '--- post the message
            MyMessenger.PostMessage(MyMessenger.Config.BreadcrumbsQ, "", "", "", nvp)

        End If

    End Sub


    ''' <summary>
    ''' Return a unique identifier
    ''' </summary>
    ''' <returns>Currently the result of a call to Guid.NewGuid</returns>
    ''' <remarks>SDCP 22-Jan-2008</remarks>
    Private Function GetNewUID() As String
        Dim sUid As String

        sUid = Guid.NewGuid().ToString
        Return sUid

    End Function


End Class

Public Module CrumblerHelper

    Public Function zzzzzCrumblerFromNVPs(ByVal NVPs As Hashtable) As Crumbler
        Dim sTrailLabel As String = ""
        Dim sTrailDescription As String = ""
        Dim sTrailUid As String = ""

        '--- try to load from NVP hashtable
        If NVPs.ContainsKey("BC_TrailLabel") Then sTrailLabel = NVPs("BC_TrailLabel").ToString
        If NVPs.ContainsKey("BC_TrailDescription") Then sTrailDescription = NVPs("BC_TrailDescription").ToString
        If NVPs.ContainsKey("BC_UID") Then sTrailUid = NVPs("BC_UID").ToString

        If (sTrailLabel <> "" And sTrailDescription <> "" And sTrailUid <> "") Then
            Return New Crumbler(sTrailLabel, sTrailDescription, sTrailUid)
        Else
            Return Nothing
        End If

    End Function

End Module

