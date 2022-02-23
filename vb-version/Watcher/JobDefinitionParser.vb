Public Class JobDefinitionParser

    Private _jobDefinitionsXml As Xml.XmlDocument

    Public Property JobDefinitionsXml() As Xml.XmlDocument
        Get
            Return _jobDefinitionsXml
        End Get
        Set(ByVal value As Xml.XmlDocument)
            _jobDefinitionsXml = value
        End Set
    End Property


    ''' <summary>
    ''' build the jobgroups and jobs
    ''' </summary>
    ''' <remarks>SDCP 26-Oct-2006</remarks>
    Public Sub Parse()

        If Not (_jobDefinitionsXml Is Nothing) Then

            '--- load in any included job definition files
            _jobDefinitionsXml.LoadXml(BootstrapIncludes(_jobDefinitionsXml.OuterXml))

            '--- expand the contextual symbols
            ExpandContextualSymbols()

            '--- do the parsing
            CreateJobGroups()

        End If

    End Sub


    

    ''' <summary>
    ''' Call the expander for contextual symbols, i.e. items that need to know their place in the Xml
    ''' e.g. JOB_GROUP_NAME
    ''' and which won't be available in the individual job's Xml
    ''' </summary>
    ''' <remarks>SDCP 06-Nov-2006</remarks>
    Private Sub ExpandContextualSymbols()

        '--- create the expander and load it with the job definition Xml
        Dim oExpander As SymbolicsExpander = New SymbolicsExpander(_jobDefinitionsXml, SymbolicsExpander.Scope.PreTrigger)
        oExpander.Expand()

        '--- get back the expanded text
        Dim sExpandedText As String = oExpander.ExpandedString

        '--- refill the DOM Doc with this string
        _jobDefinitionsXml.LoadXml(sExpandedText)

    End Sub


    ''' <summary>
    ''' creates the jobgroups and populates with their jobs
    ''' </summary>
    ''' <remarks>SDCP 26-Oct-2006</remarks>
    Private Sub CreateJobGroups()

        '--- create new job group collection
        JobGroupList = New List(Of JobGroup)

        '--- get all jobgroups from the def
        Dim nlJobGroups As Xml.XmlNodeList = _jobDefinitionsXml.DocumentElement.SelectNodes("descendant::jobgroup")

        '--- create the job groups
        For Each nJobGroup As Xml.XmlNode In nlJobGroups

            Dim jg As New JobGroup
            jg.Name = nJobGroup.Attributes("name").Value

            '--- create the Jobs
            For Each nJob As Xml.XmlNode In nJobGroup.SelectNodes("descendant::job")

                Dim j As New Job

                Try

                    '--- extract some properties
                    Try
                        j.Name = nJob.Attributes("name").Value
                        j.JobXml = nJob
                    Catch
                        MyMessenger.AtomInfo += "<error>Job has no name [A job must have a name attribute]</error>"
                        MyMessenger.AtomInfo += "<jobdefcontents>" & nJob.OuterXml & "</jobdefcontents>"
                        MyMessenger.PostMessage(MyMessenger.Config.StatusQ, "", "", "", GetWatcherErrorNVPs)

                        '--- fail this jobdef, but allow others to be created
                        Exit For

                    End Try

                    '--- extract local variables as required
                    j.ExtractLocalVariables()

                    '--- create the trigger objects
                    j.BuildTrigger()

                Catch ex As Exception

                    MyMessenger.AtomInfo += "<error>JobDef failed to build correctly</error>"
                    MyMessenger.AtomInfo += "<jobdefcontents>" & nJob.OuterXml & "</jobdefcontents>"
                    MyMessenger.PostMessage(MyMessenger.Config.StatusQ, "", "", "", GetWatcherErrorNVPs)

                    '--- fail this jobdef, but allow others to be created
                    Exit For

                End Try

                '--- add to the jobgroup
                jg.Add(j)

            Next

            '--- add the job group to the global collection
            JobGroupList.Add(jg)

        Next

    End Sub

End Class
