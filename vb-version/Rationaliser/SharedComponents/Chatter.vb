Imports OTIS

Public Class Chatter

    Private _ChatEMSServer As String
    Private _ChatEMSUserName As String
    Private _ChatEMSPassword As String
    Private _ChatQueue As String

    Private _Otis As OTIS.OTIS

    Private _Recipients As Hashtable
    Private _RunLabel As String

    Dim MyChatMessenger As Messenger

    Public Enum DestType
        Channel = 1
        Individual = 2
    End Enum

    Public ReadOnly Property RunLabel() As String
        Get
            Return _RunLabel
        End Get
    End Property

    ''' <summary>
    ''' Read tib settings for chatting from the app config.  Setup OTIS
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()

        'Environment-specific app config already expanded.  Use it
        Dim xDoc As Xml.XmlDocument = MyMessenger.Config.DomDoc

        'Now pick out settings for chatting
        Dim xNode As Xml.XmlNode
        xNode = xDoc.SelectSingleNode("//RPMConfig/chat/server")
        If xNode Is Nothing Then
            MyMessenger.Log(Messenger.LogType.ErrorType, "Unable to chat - couldn't find chat ems server setting")
            Exit Sub
        Else
            _ChatEMSServer = xNode.InnerText
        End If

        xNode = xDoc.SelectSingleNode("//RPMConfig/chat/username")
        If xNode Is Nothing Then
            MyMessenger.Log(Messenger.LogType.ErrorType, "Unable to chat - couldn't find chat ems username")
            Exit Sub
        Else
            _ChatEMSUserName = xNode.InnerText
        End If

        xNode = xDoc.SelectSingleNode("//RPMConfig/chat/password")
        If xNode Is Nothing Then
            MyMessenger.Log(Messenger.LogType.ErrorType, "Unable to chat - couldn't find chat ems password")
            Exit Sub
        Else
            _ChatEMSPassword = xNode.InnerText
        End If

        xNode = xDoc.SelectSingleNode("//RPMConfig/chat/queue[@name='Grapevine']")
        If xNode Is Nothing Then
            MyMessenger.Log(Messenger.LogType.ErrorType, "Unable to chat - couldn't find chat queue name")
            Exit Sub
        Else
            _ChatQueue = xNode.InnerText
        End If

        _Otis = New OTIS.OTIS(_ChatEMSServer, _ChatEMSUserName, _ChatEMSPassword)
        _Recipients = New Hashtable

    End Sub

    ''' <summary>
    ''' Take a list of recipients from a job def
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ReadChatSettings()

        Dim xNodeList As Xml.XmlNodeList, xNode As Xml.XmlNode
        xNodeList = MyMessenger.Config.DomDoc.SelectNodes("//coreapps/application[@name='" & sTHIS_APP & "']/chat/who")
        For Each xNode In xNodeList
            _Recipients.Add(_Recipients.Count + 1, xNode.InnerText)
        Next

        _RunLabel = sTHIS_APP

    End Sub

    ''' <summary>
    ''' Take a list of recipients from a job def
    ''' </summary>
    ''' <param name="JobDef"></param>
    ''' <remarks></remarks>
    Public Sub ReadChatSettings(ByVal JobDef As String)

        Dim xDoc As Xml.XmlDocument = New Xml.XmlDocument
        xDoc.LoadXml(JobDef)

        Dim xNodeList As Xml.XmlNodeList, xNode As Xml.XmlNode
        xNodeList = xDoc.DocumentElement.SelectNodes("atom/chat/who")
        For Each xNode In xNodeList
            _Recipients.Add(_Recipients.Count + 1, xNode.InnerText)
        Next

        xNode = xDoc.DocumentElement.SelectSingleNode("atom/chat/label")
        If Not xNode Is Nothing Then
            _RunLabel = xNode.InnerText
        Else
            _RunLabel = sTHIS_APP & " running: " & Format(Now, "dd/MM/yy HH:mm:ss")
        End If

    End Sub

    ''' <summary>
    ''' Post a message onto the chat queue.
    ''' </summary>
    ''' <param name="Destination">The person or channel you're chatting to</param>
    ''' <param name="DestinationType">Either a person or a channel</param>
    ''' <param name="MessageBody">the message you want to send</param>
    ''' <remarks></remarks>
    Public Sub Chat(ByVal Destination As String, ByVal DestinationType As DestType, ByVal MessageBody As String)

        'Hashtable of NVPs
        Dim ChatMessage As Hashtable = New Hashtable
        ChatMessage.Add("destination", Destination)
        ChatMessage.Add("messagebody", MessageBody)

        If DestinationType = DestType.Channel Then
            ChatMessage.Add("destinationtype", "channel")
        ElseIf DestinationType = DestType.Individual Then
            ChatMessage.Add("destinationtype", "individual")
        End If

        'Send the message
        Try
            Dim gunMessage As GunMessage = _Otis.PublishMessage(_ChatQueue, ChatMessage)
        Catch ex As Exception
        End Try

    End Sub

    ''' <summary>
    ''' OVERLOAD
    ''' Just supply the message body.  Loop through all the recipients in the hashtable and send same message to 
    ''' them all
    ''' </summary>
    ''' <param name="MessageBody">message to be sent to all recipients</param>
    ''' <remarks></remarks>
    Public Sub Chat(ByVal MessageBody As String)

        For Each de As DictionaryEntry In _Recipients

            'Hashtable of NVPs
            Dim ChatMessage As Hashtable = New Hashtable
            Dim Destination As String

            Try
                Destination = de.Value.ToString

                'It's a channel if destination starts with a #
                If Destination.Substring(1, 1) = "#" Then
                    ChatMessage.Add("destinationtype", "channel")
                    Destination = Destination.Substring(2)
                Else
                    ChatMessage.Add("destinationtype", "individual")
                End If

                ChatMessage.Add("destination", Destination)
                ChatMessage.Add("messagebody", MessageBody)

            Catch ex As Exception

            End Try

            'Send the message
            Try
                Dim gunMessage As GunMessage = _Otis.PublishMessage(_ChatQueue, ChatMessage)
            Catch ex As Exception
            End Try

        Next

    End Sub

End Class
