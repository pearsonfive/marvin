Imports System.Xml
Imports System.IO


Module WhenFinished

    ''' <summary>
    ''' Work out if anything needs to happen after atom process is finished and do it.
    ''' </summary>
    ''' <param name="XMLPayload"></param>
    ''' <returns>pass through boolean from MessageActions or FileActions</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>

    Function DoWhenFinishedActions(ByVal XMLPayload As String, ByVal MyMessenger As Messenger) As Boolean
        Dim fResult As Boolean

        'Load the payload into a dom doc
        Dim xDoc As New Xml.XmlDocument, xNodeList As Xml.XmlNodeList
        xDoc.LoadXml(XMLPayload)

        'See if a "whenfinished" node exists
        Dim xPath As String = "//job/whenfinished"

        '--- default return value is True, i.e. if there is no <whenfinished> node
        fResult = True

        xNodeList = xDoc.DocumentElement.SelectNodes(xPath)
        For Each xNode As Xml.XmlNode In xNodeList

            fResult = MessageActions(xNode, XMLPayload, MyMessenger)

            If fResult = False Then Exit For

        Next

        Return fResult

    End Function

    ''' <summary>
    ''' If the action to take was defined as a message, determine message settings and post accordingly
    ''' </summary>
    ''' <param name="messageActionsNode"></param>
    ''' <param name="XMLPayload"></param>
    ''' <returns>Boolean</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>

    Function MessageActions(ByVal messageActionsNode As Xml.XmlNode, _
                                        ByVal XMLPayload As String, ByVal MyMessenger As Messenger) As Boolean

        Dim sExpiryUnit As String = "second"
        Dim lExpiryAmount As Integer = 0
        Dim lExpirySeconds As Integer = 0
        Dim xPath As String = ""
        Dim xNode As Xml.XmlNode

        Try

            'Get the queue name
            Dim sQueueName As String = messageActionsNode.SelectSingleNode("queue").InnerText

            '--- extract expiry info
            xPath = "@expiryunit"
            xNode = messageActionsNode.SelectSingleNode(xPath)
            If Not xNode Is Nothing Then
                sExpiryUnit = xNode.InnerText
            End If
            xPath = "@expiryamount"
            xNode = messageActionsNode.SelectSingleNode(xPath)
            If Not xNode Is Nothing Then
                lExpiryAmount = CInt(xNode.InnerText)
            End If

            If lExpiryAmount <> 0 Then
                lExpirySeconds = CInt((GetAdjustedDate(Now, sExpiryUnit, lExpiryAmount) - Now).TotalSeconds)
            End If

            Dim nvp As Hashtable = New Hashtable

            'Write all name/values into a hashtable
            Dim nameValuePairsNodeList As Xml.XmlNodeList = messageActionsNode.SelectNodes("messagecontents/namevaluepair")
            For Each xChNode As Xml.XmlNode In nameValuePairsNodeList
                nvp.Add(xChNode.SelectSingleNode("name").InnerText, xChNode.SelectSingleNode("value").InnerText)
            Next

            '--- for free add the original payload too
            nvp.Add("originalpayload", XMLPayload)




            'Post with empty hash/payload/jobname settings - these are included in the nvp hashtable
            MyMessenger.PostMessage(sQueueName, "", "", "", nvp, "", "", lExpirySeconds)

            '--- drop crumb
            MyCrumbler.DropCrumb("DoWhenFinished", GetWhenFinishedInfo(messageActionsNode))

            Return True

        Catch ex As Exception
            Dim sMessage As String = "#Error: Exception in WhenFinished/MessageActions" & vbCr
            sMessage = sMessage & "<ExceptionMessage>" & ex.Message & "</ExceptionMessage>"
            sMessage = sMessage & "<ExceptionStack>" & ex.StackTrace & "</ExceptionStack>"
            sMessage = sMessage & "<ProblemNodeText>" & messageActionsNode.OuterXml & "</ProblemNodeText>"
            MyMessenger.Log(Messenger.LogType.ErrorType, sMessage, MySwitches("hash"))
            Return False
        End Try

    End Function


    ''' <summary>
    ''' creates the crumbler info node for a WhenFnished action
    ''' </summary>
    ''' <param name="WhenFinishedNode"></param>
    ''' <returns></returns>
    ''' <remarks>SDCP 31-Jan-2008</remarks>
    Private Function GetWhenFinishedInfo(ByVal WhenFinishedNode As xmlnode) As xmlnode
        Dim xDoc As XmlDocument
        

        '--- create the doc
        xDoc = New XmlDocument
        xDoc.LoadXml("<info/>")

        '--- just write in the WhenFinished node
        xDoc.DocumentElement.InnerXml = WhenFinishedNode.OuterXml

        '--- return the data
        Return xDoc.DocumentElement

    End Function


End Module
