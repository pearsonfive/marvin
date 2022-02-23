Imports OTIS
Imports System.Xml
Imports System.Runtime.InteropServices
Imports System.Net.Mail


Public Class Messenger

    Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As IntPtr) As Integer
    Private Declare Unicode Function NetRemoteTOD Lib "netapi32" (<MarshalAs(UnmanagedType.LPWStr)> ByVal ServerName As String, _
        ByRef BufferPtr As IntPtr) As Integer

    Public AtomInfo As String

    Private _Otis As OTIS.OTIS

    Private _EMSServer As String
    Private _EMSUserName As String
    Private _EMSPassword As String

    Private _MachineName As String
    Private _CallingProcess As String
    Private _QAllTriggered As String
    Private _QJobsToDo As String
    Private _QJobsWorking As String
    Private _QJobsDone As String
    Private _QStatus As String

    Private _Config As Config

    Public Enum LogType
        StatusType = 1
        ErrorType = 2
    End Enum

    Structure TIME_OF_DAY_INFO
        Dim tod_elapsedt As Integer
        Dim tod_msecs As Integer
        Dim tod_hours As Integer
        Dim tod_mins As Integer
        Dim tod_secs As Integer
        Dim tod_hunds As Integer
        Dim tod_timezone As Integer
        Dim tod_tinterval As Integer
        Dim tod_day As Integer
        Dim tod_month As Integer
        Dim tod_year As Integer
        Dim tod_weekday As Integer
    End Structure

    ''' <summary>
    ''' Instantiate the messenger.  Instantiate a config, which supplies server/username/password
    ''' necessary for messaging.  Expose this so can be used subsequently in code.
    ''' </summary>
    ''' <param name="RPMConfigFile">Full path to the RPM config file</param>
    ''' <param name="ThisAppName">Name of the application we're running in</param>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Sub New(ByVal RPMConfigFile As String, ByVal ThisAppName As String)

        Dim oConfig As Config = New Config(RPMConfigFile, ThisAppName)
        _Config = oConfig
        _EMSServer = oConfig.EmsServer
        _EMSUserName = oConfig.EmsUsername
        _EMSPassword = oConfig.EmsPassword

        _Otis = New OTIS.OTIS(_EMSServer, _EMSUserName, _EMSPassword)

        _MachineName = System.Environment.MachineName
        _CallingProcess = ThisAppName

        _QAllTriggered = oConfig.AllTriggeredJobsQ
        _QJobsToDo = oConfig.JobsToDoQ
        _QJobsWorking = oConfig.JobsInProgressQ
        _QJobsDone = oConfig.JobsDoneQ
        _QStatus = oConfig.StatusQ

    End Sub

    ''' <summary>
    ''' Post a message onto a specified queue.  Also require hash, payload, jobname and hashtable of
    ''' name/value pairs (can be empty).  Job group allows users to view groups of jobs (via a message 
    ''' selector query)
    ''' </summary>
    ''' <param name="QueueName">valid queue from the rpm config</param>
    ''' <param name="Hash"></param>
    ''' <param name="Payload"></param>
    ''' <param name="JobName"></param>
    ''' <param name="NameValuePairs"></param>
    ''' <param name="JobGroup"></param>
    ''' <returns>Message ID if successful; error message if not</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function PostMessage(ByVal QueueName As String, _
                                            ByVal Hash As String, _
                                            ByVal Payload As String, _
                                            ByVal JobName As String, _
                                            ByVal NameValuePairs As Hashtable, _
                                            Optional ByVal JobGroup As String = "", _
                                            Optional ByVal StatusMessages As String = "", _
                                            Optional ByVal ExpirySeconds As Integer = 0) As String

        'Just in case they've put any of these in the NVP, ignore any errors
        On Error Resume Next

        '--- SDCP 02-Jan-2007
        '--- ensure all these are valid strings
        NameValuePairs.Add("hash", IIf(Hash = "", MySwitches("hash"), Hash))
        NameValuePairs.Add("payload", IIf(Payload Is Nothing, "", Payload))
        NameValuePairs.Add("jobname", IIf(JobName Is Nothing, "", JobName))
        NameValuePairs.Add("jobgroup", IIf(JobGroup Is Nothing, "", JobGroup))
        NameValuePairs.Add("queuename", IIf(QueueName Is Nothing, "", QueueName))


        'Pick out existing audit trail
        Dim sAudit As String = "", iAuditCount As Integer
        For Each de As DictionaryEntry In NameValuePairs
            If de.Key.ToString = "audit" Then
                sAudit = de.Value.ToString
            ElseIf de.Key.ToString = "auditcount" Then
                iAuditCount = CType(de.Value, Integer)
            End If
        Next

        If sAudit = "" Then
            sAudit = "<audittrail></audittrail>"
            iAuditCount = 0
        Else
            NameValuePairs.Remove("audit")
            NameValuePairs.Remove("auditcount")
        End If

        'Audit info for this point
        Dim sThisAudit As String = "</audit></audittrail>"
        sThisAudit = "<hash>" & Hash & "</hash>" & sThisAudit
        sThisAudit = "<timestamp>" & Format(Now, "yyyyMMdd HH:mm:ss") & "</timestamp>" & sThisAudit
        sThisAudit = "<environment>" & MySwitches("environment") & "</environment>" & sThisAudit
        sThisAudit = "<queuename>" & QueueName & "</queuename>" & sThisAudit
        sThisAudit = "<application>" & sTHIS_APP & "</application>" & sThisAudit
        If AtomInfo <> "" Then
            sThisAudit = "<atominfo>" & AtomInfo & "</atominfo>" & sThisAudit
        End If
        sThisAudit = "<username>" & UserName() & "</username>" & sThisAudit
        sThisAudit = "<machinename>" & MachineName() & "</machinename>" & sThisAudit
        If StatusMessages <> "" Then
            sThisAudit = "<statusmessages>" & StatusMessages & "</statusmessages>" & sThisAudit
        End If
        sThisAudit = "<audit>" & sThisAudit

        'Append to existing audit info
        sAudit = Replace(sAudit, "</audittrail>", sThisAudit)
        NameValuePairs.Add("audit", sAudit)
        NameValuePairs.Add("auditcount", iAuditCount + 1)

        On Error GoTo ErrHandler
        Dim gunMessage As GunMessage = _Otis.PublishMessage(QueueName, NameValuePairs, ExpirySeconds)
        Return gunMessage.MessageID
        Exit Function

ErrHandler:
        Return "#Error: " & Err.GetException.StackTrace

    End Function

    Public Function PostMessage(ByVal QueueName As String, _
                                            ByVal MyMessage As GunMessage) As String

        '--- de-serialise the message 
        Dim nvp As Hashtable = MyMessage.Pairs

        'For Each de As DictionaryEntry In MyMessage.Names
        '    nvp.Add(de.Value, MyMessage.GetNamedValue(de.Value.ToString))
        'Next

        '--- post the message
        Return PostMessage(QueueName, MyMessage.Hash, MyMessage.Payload, MyMessage.JobName, nvp)


    End Function

    ''' <summary>
    ''' Consume message from a specified queue by selector
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="MessageSelector"></param>
    ''' <returns>A GunMessage if successful; otherwise nothing</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function ConsumeMessageBySelector(ByVal QueueName As String, _
                                            ByVal MessageSelector As String) As GunMessage

        Dim gunMessage As GunMessage

        Try

            'xHIGH: in MyMessenger::ConsumeMessageBySelector() need to make the timeout flexible
            gunMessage = _Otis.ConsumeMessage(QueueName, 1000, MessageSelector)

            Return gunMessage

        Catch ex As Exception
            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' Consume either a specific message (given a message ID) or any message from a specified queue.  
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="MessageID"></param>
    ''' <returns>A GunMessage if successful; otherwise nothing</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function ConsumeMessage(ByVal QueueName As String, _
                                            Optional ByVal MessageID As String = "") As GunMessage

        Dim gunMessage As GunMessage

        Try


            'xHIGH: in MyMessenger::ConsumeMessage() need to make the timeout flexible
            If MessageID = "" Then
                gunMessage = _Otis.ConsumeMessage(QueueName, 1000)
            Else
                gunMessage = _Otis.ConsumeMessageByID(QueueName, 1000, MessageID)
            End If


            Return gunMessage

        Catch ex As Exception
            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' Consume either a specific message (given a message ID) or any message from a specified queue.  
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="MessageID"></param>
    ''' <returns>A GunMessage if successful; otherwise nothing</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function ConsumeMessage(ByVal QueueName As String, ByVal MessageID As String, ByVal TimeOut As Long) As GunMessage

        Dim gunMessage As GunMessage

        Try

            If MessageID = "" Then
                gunMessage = _Otis.ConsumeMessage(QueueName, TimeOut)
            Else
                gunMessage = _Otis.ConsumeMessageByID(QueueName, TimeOut, MessageID)
            End If


            Return gunMessage

        Catch ex As Exception
            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' Given a queue and a query string (message selector), return all messages that conform to
    ''' that query
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="MessageSelector">see TIB documentation for syntax (similar to SQL though)</param>
    ''' <returns>A hashtable containing GunMessages</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function SnapshotQueue(ByVal QueueName As String, _
                                    Optional ByVal MessageSelector As String = "", Optional ByVal ReturnIDsOnly As Boolean = False) As Hashtable
        Dim ht As New Hashtable

        Try

            ht = _Otis.QueueSnapshot(QueueName.ToString, MessageSelector, ReturnIDsOnly)

        Catch ex As Exception

        End Try

        Return ht

    End Function

    ''' <summary>
    ''' Given a queue and a query string (message selector), return count of messages that conform to
    ''' that query
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="MessageSelector">see TIB documentation for syntax (similar to SQL though)</param>
    ''' <returns>Integer</returns>
    ''' <remarks>SDCP 12-Dec-2006</remarks>
    Public Function QueueCount(ByVal QueueName As String, _
                                    Optional ByVal MessageSelector As String = "") As Integer

        Dim lCount As Integer

        Try

            lCount = _Otis.QueueCount(QueueName.ToString, MessageSelector)

        Catch ex As Exception

        End Try

        Return lCount

    End Function

    ''' <summary>
    ''' Given a queue and a query string (message selector), return all messages that conform to
    ''' that query
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="MessageSelector">see TIB documentation for syntax (similar to SQL though)</param>
    ''' <returns>A hashtable containing GunMessages</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function SnapshotQueue_Ordered(ByVal QueueName As String, _
                                    Optional ByVal MessageSelector As String = "") As Collection
        Dim col As New Collection

        Try

            col = _Otis.QueueSnapshot_Ordered(QueueName.ToString, MessageSelector)

        Catch ex As Exception

        End Try

        Return col

    End Function


    ''' <summary>
    ''' Given a queue and a query string (message selector), return all messages that conform to
    ''' that query
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="MessageSelector">see TIB documentation for syntax (similar to SQL though)</param>
    ''' <returns>A hashtable containing GunMessages</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function SnapshotQueue_Ordered_AsList(ByVal QueueName As String, _
                                    Optional ByVal MessageSelector As String = "") As List(Of GunMessage)

        Dim lstGunMsg As New List(Of GunMessage)

        Try

            lstGunMsg = _Otis.QueueSnapshot_Ordered_AsList(QueueName.ToString, MessageSelector)

        Catch ex As Exception

        End Try

        Return lstGunMsg

    End Function


    ''' <summary>
    ''' Identical to a Post, but provides a default info set.  Log can be a Status or Error type.
    ''' Before posting, check for another log message with same hash, and append to this rather than posting many 
    ''' messages about the same job
    ''' </summary>
    ''' <param name="LogType"></param>
    ''' <param name="LogMessage"></param>
    ''' <param name="Hash"></param>
    ''' <returns></returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Public Function Log(ByVal LogType As LogType, _
                                ByVal LogMessage As String, _
                                Optional ByVal Hash As String = "") As String

        Try
            Dim MsgSelector As String = "hash=" & Qt(Hash)
            Dim StatusMessage As GunMessage = MyMessenger.ConsumeMessageBySelector(MyMessenger.Config.StatusQ, MsgSelector)

            'If there are no messages about this hash on the queue, set up a new audit message
            Dim Audit As String
            If StatusMessage Is Nothing Then
                Audit = "<audit></audit>"
            Else
                Try
                    Audit = StatusMessage.Pairs("audit").ToString
                Catch
                    Audit = "<audit></audit>"
                End Try
            End If

            Dim ThisAudit As String = "<logmessage>"
            ThisAudit += "<type>" & LogType.ToString & "</type>"
            ThisAudit += "<message>" & LogMessage & "</message>"
            ThisAudit += "<machinename>" & _MachineName & "</machinename>"
            ThisAudit += "<callingapp>" & _CallingProcess & "</callingapp>"
            ThisAudit += "<hash>" & Hash & "</hash>"
            ThisAudit += "<timestamp>" & Now.ToShortDateString & " " & Now.ToShortTimeString & "</timestamp>"
            If AtomInfo <> "" Then
                ThisAudit = "<atominfo>" & AtomInfo & "</atominfo>" & ThisAudit
            End If
            ThisAudit += "</logmessage>"

            Audit = Audit.Replace("</audit>", ThisAudit) & "</audit>"

            Dim ht As New Hashtable
            ht.Add("hash", Hash)
            ht.Add("audit", Audit)

            Dim _GunMessage As GunMessage = _Otis.PublishMessage(_QStatus, ht)
            Return _GunMessage.MessageID

        Catch ex As Exception

            Return "#Err: " & ex.Message

        End Try

    End Function

    Public ReadOnly Property Config() As Config
        Get
            Return _Config
        End Get
    End Property

    ''' <summary>
    ''' Tests for the existence of a message with the given hash on the selected queue
    ''' </summary>
    ''' <param name="Hash">The hash to be tested for</param>
    ''' <param name="QueueName">The queue name to look on</param>
    ''' <returns>True if found, False if not</returns>
    ''' <remarks>SDCP 10-Oct-2006</remarks>
    Public Function HashExistsOnQueue(ByVal Hash As String, ByVal QueueName As String) As Boolean
        Dim htMessages As Hashtable
        Dim sSelector As String


        '--- create the message selector
        sSelector = "hash=" & Qt(Hash)

        '--- browse OTIS for messages with matching hash
        htMessages = _Otis.QueueSnapshot(QueueName, sSelector)

        '--- return the result
        If htMessages IsNot Nothing Then
            Return (htMessages.Count > 0)
        Else
            Return False
        End If

    End Function


    ''' <summary>
    ''' Creates a disconnected GunMessage, e.g. when you need Gun Message capabilities 
    ''' but don't want to write the message to TIB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>SDCP 19-Jun-2007</remarks>
    Public Function CreateDisconnectedMessage() As GunMessage
        'Dim gq As GunQueue
        Dim gmsg As GunMessage = Nothing

        'Try

        '    gq = _Otis.CreateQueue()
        '    gmsg = gq.CreateMapMessage()
        '    gq = Nothing
        '    Return gmsg

        'Finally



        'End Try

        Try

            gmsg = _Otis.PublishMessage("simon_Trash", New Hashtable(), 1)

        Catch ex As Exception

            Stop

        End Try

        Return gmsg

    End Function


    ''' <summary>
    ''' Purge a queue with optional MessageSelector
    ''' </summary>
    ''' <param name="QueueName"></param>
    ''' <param name="MessageSelector"></param>
    ''' <remarks>SDCP 03-Nov-2006</remarks>
    Public Sub PurgeQueue(ByVal QueueName As String, Optional ByVal MessageSelector As String = "")
        Dim gunMessage As GunMessage

        Try

            Do

                gunMessage = Me.ConsumeMessageBySelector(QueueName, MessageSelector)

            Loop Until gunMessage.MessageID = ""

        Catch ex As Exception
        End Try

    End Sub

    Function BroadcastMessage(ByVal TopicName As String, ByVal MessageText As String) As GunMessage

        Return _Otis.BroadcastMessage(TopicName, MessageText)

    End Function


    Public Function ActiveTIBServer() As String
        Return _Otis.ActiveTIBServer
    End Function

    ''' <summary>
    ''' In the absence of a good call to tib to establish the time, request the time from the active tib server box - have to
    ''' strip out any port info
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TIBTime() As Long

        Dim ActiveServer As String = ActiveTIBServer()
        Dim iPos As Integer = InStr(ActiveServer, ":")
        If iPos > 0 Then ActiveServer = Left(ActiveServer, iPos - 1)

        Dim ret As Double = GetNetRemoteTOD(ActiveServer).Ticks
        ret = ret - DateSerial(1970, 1, 1).Ticks
        ret = Math.Round(ret / TimeSpan.TicksPerMillisecond, 0)

        Return CType(ret, Long)

    End Function

    ''' <summary>
    ''' The idea of this was to call tib in a light touch manner, and if it responds it must be up.  TibTime causing problems, so 
    ''' removed 11/5/07.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Martin Leyland, May 2007</remarks>
    Public Function TibIsUp() As Boolean
        Try
            If IsNumeric(_Otis.TIBTime()) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    ''' <summary>
    ''' http://www.codeproject.com/vb/net/NetRemoteTOD.asp
    ''' Get the time of day from a remote box.
    ''' </summary>
    ''' <param name="strServerName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNetRemoteTOD(ByVal strServerName As String) As Date

        Try
            Dim iRet As Integer
            Dim ptodi As IntPtr
            Dim todi As TIME_OF_DAY_INFO
            Dim dDate As Date
            strServerName = strServerName & vbNullChar
            iRet = NetRemoteTOD(strServerName, ptodi)
            If iRet = 0 Then
                todi = CType(Marshal.PtrToStructure(ptodi, GetType(TIME_OF_DAY_INFO)), _
                  TIME_OF_DAY_INFO)
                NetApiBufferFree(ptodi)
                dDate = DateSerial(todi.tod_year, todi.tod_month, todi.tod_day).AddHours(todi.tod_hours)
                dDate = dDate.AddMinutes(todi.tod_mins)
                dDate = dDate.AddSeconds(todi.tod_secs)
                GetNetRemoteTOD = dDate
            Else
                Return Now
            End If
        Catch
            Return Now
        End Try
    End Function


    Public Function SendEmail(ByVal From As String, ByVal Target As String, ByVal Subject As String, ByVal Body As String) As Boolean
        Const SMTP_SERVER As String = "mailhost"
        Const SMTP_PORT As Long = 25

        Dim message As New MailMessage(From, Target)
        message.Subject = Subject
        message.Body = Body

        Dim client As New SmtpClient(SMTP_SERVER, SMTP_PORT)
        client.UseDefaultCredentials = True
        client.Send(message)

    End Function

    Public Function SendEmail(ByVal From As String, ByVal Target As String, ByVal Subject As String, ByVal Body As String, ByVal MyAttachments As Collection) As Boolean
        Const SMTP_SERVER As String = "mailhost"
        Const SMTP_PORT As Long = 25

        Dim message As New MailMessage(From, Target)
        message.Subject = Subject
        message.Body = Body
        For Each ma As Attachment In MyAttachments
            message.Attachments.Add(ma)
        Next ma

        Dim client As New SmtpClient(SMTP_SERVER, SMTP_PORT)
        client.UseDefaultCredentials = True
        client.Send(message)

    End Function


    Public Function CreateTopic() As GunTopic
        Return _Otis.CreateTopic
    End Function

End Class
