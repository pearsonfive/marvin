Public Class Trigger

    Private _triggerInfoXml As Xml.XmlNode
    Private _expandedTriggerInfoXml As Xml.XmlNode
    Private _hash As String
    Private _watcherType As String
    Private _name As String
    Private _timestamp As Date
    Private _uniqueID As String


    Public Property TriggerInfoXml() As Xml.XmlNode
        Get
            Return _triggerInfoXml
        End Get
        Friend Set(ByVal value As Xml.XmlNode)
            _triggerInfoXml = value
        End Set
    End Property


    Public Property ExpandedTriggerInfoXml() As Xml.XmlNode
        Get
            Return _expandedTriggerInfoXml
        End Get
        Friend Set(ByVal value As Xml.XmlNode)
            _expandedTriggerInfoXml = value
        End Set
    End Property


    Public Property UniqueID() As String
        Get
            Return _uniqueID
        End Get
        Friend Set(ByVal value As String)
            _uniqueID = value
        End Set
    End Property


    Public Property Timestamp() As Date
        Get
            Return _timestamp
        End Get
        Set(ByVal value As Date)
            _timestamp = value
        End Set
    End Property


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


End Class
