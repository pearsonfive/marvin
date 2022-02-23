Imports OTIS
Imports System.Management
Imports System.Diagnostics

Module HelperFunctions

#Region "API Calls"
    ' standard API declarations for INI access
    ' changing only "As Long" to "As Int32" (As Integer would work also)
    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Int32

    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Int32, _
    ByVal lpFileName As String) As Int32
#End Region



    ''' <summary>
    ''' Quotes the string with single quotes
    ''' </summary>
    ''' <param name="Str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function Qt(ByVal Str As String) As String
        Return Chr(39) & Str & Chr(39)
    End Function

    ''' <summary>
    ''' Quotes the string with double quotes
    ''' </summary>
    ''' <param name="Str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function DQt(ByVal Str As String) As String
        Return Chr(34) & Str & Chr(34)
    End Function

    ''' <summary>
    ''' Return a sorted list of nodes based on sorting by an attribute
    ''' </summary>
    ''' <param name="NodeList">Nodelist that needs sorting</param>
    ''' <param name="KeyAttribute">Name of attribute to use as key</param>
    ''' <param name="TypeCode">The typecode of the attribute, e.g. "String", "Int32"</param>
    ''' <returns>A SortedList</returns>
    ''' <remarks>SDCP 11-Oct-2006</remarks>

    Public Function SortNodeList(ByVal NodeList As Xml.XmlNodeList, ByVal KeyAttribute As String, Optional ByVal TypeCode As TypeCode = TypeCode.String) As SortedList
        Dim list As New SortedList

        For Each el As Xml.XmlElement In NodeList

            '--- select the user-specified attribute as the key:
            Dim key As String

            Try

                '--- extract the Key
                key = CType(el, Xml.XmlElement).Attributes(KeyAttribute).Value

                '--- add to the sorted list using the key (in correct format)
                If TypeCode = VariantType.String Then

                    '--- key is type String (default behavior)
                    list.Add(key, el)

                Else

                    '--- convert value to type specified by caller:
                    list.Add(Convert.ChangeType(key, TypeCode), el)

                End If

            Catch ex As Exception '--- if there is an Exception then don't add
            End Try

        Next

        '--- return the sorted list
        Return list

    End Function

    ''' <summary>
    ''' Return the name of the logged in user
    ''' </summary>
    ''' <returns>string</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Function UserName() As String
        Return System.Environment.UserName
    End Function


    ''' <summary>
    ''' Return the machine name
    ''' </summary>
    ''' <returns>string</returns>
    ''' <remarks>Martin Leyland, October 2006</remarks>
    Function MachineName() As String
        Return System.Environment.MachineName.ToUpper
    End Function


    ''' <summary>
    ''' returns the current process' PID
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CurrentPID() As Integer
        Return Process.GetCurrentProcess.Id
    End Function



    ''' <summary>
    ''' returns a date adjusted appropriately
    ''' </summary>
    ''' <param name="StartDate">the date to adjust</param>
    ''' <param name="OffsetByUnit">e.g. hour, day, week, month</param>
    ''' <param name="OffsetByAmount">an integer amount of units to offset by</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAdjustedDate(ByVal StartDate As Date, ByVal OffsetByUnit As String, ByVal OffsetByAmount As Integer) As Date
        Dim returnValue As Date

        Select Case OffsetByUnit.ToLower()

            Case "second"
                returnValue = StartDate.AddSeconds(OffsetByAmount)
            Case "minute"
                returnValue = StartDate.AddMinutes(OffsetByAmount)
            Case "hour"
                returnValue = StartDate.AddHours(OffsetByAmount)
            Case "day"
                returnValue = StartDate.AddDays(OffsetByAmount)
            Case "month"
                returnValue = StartDate.AddMonths(OffsetByAmount)
            Case "year"
                returnValue = StartDate.AddYears(OffsetByAmount)
        End Select

        Return returnValue

    End Function

    ''' <summary>
    ''' Loops through all processes, count each of these processes
    ''' </summary>
    ''' <param name="ProcessName">Process to count</param>
    ''' <returns>Integer</returns>
    ''' <remarks>Martin Leyland, December 2006</remarks>
    Public Function CountProcesses(ByVal ProcessName As String) As Integer

        Dim p() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
        For i As Integer = 0 To p.GetUpperBound(0)
            If LCase(p(i).ProcessName) = LCase(ProcessName) Then
                CountProcesses += 1
            End If
        Next

    End Function



    Public Function CreateUniqueID() As String
        Randomize()
        Return LCase(Format(Now, "ddMMMyy") & "_" & Format(Now, "HHmmss") & "_" & Format(Int(Rnd() * 9999999), "0000000"))
        'Return LCase(Format(Now, "ddmmmyy") & "_" & Format(Now, "hhmmss") & "_" & Now.Ticks)
    End Function


    ''' <summary>
    ''' verify and bootstrap in any "includes" in the input doc (recursively)
    ''' </summary>
    ''' <param name="XMLString">The input xml string</param>
    ''' <returns></returns>
    ''' <remarks>SDCP 15-Nov-2006</remarks>
    Public Function BootstrapIncludes(ByVal XMLString As String) As String

        Dim expandedJobDef As String = XMLString.Replace("#[ENVIRONMENT]#", MySwitches("environment"))

        '--- create the Xml doc
        Dim xJobDefs As New Xml.XmlDocument
        xJobDefs.LoadXml(expandedJobDef)

        For Each xInclude As Xml.XmlNode In xJobDefs.SelectNodes("//include")

            '--- load the<include> into a Dom doc
            Dim xDoc As Xml.XmlDocument = New Xml.XmlDocument

            Try


                Try

                    '--- load the referenced file
                    xDoc.Load(xInclude.Attributes("href").InnerText)

                Catch ex As Xml.XmlException

                    MyMessenger.AtomInfo += "<error>Failed to parse [include] file</error>"
                    MyMessenger.AtomInfo += "<errormessage>" & CDataBlock(ex.Message) & "</errormessage>"
                    MyMessenger.AtomInfo += "<includeaddress>" & CDataBlock(xInclude.Attributes("href").InnerText) & "</includeaddress>"
                    MyMessenger.PostMessage(MyMessenger.Config.StatusQ, "", "", "", New Hashtable)
                    MyMessenger.SendEmail("DL-RatesDeskDev@ubs.com", "DL-RatesDeskDev@ubs.com", "Watcher failure on " & My.Computer.Name, "Failed to parse [include] file" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace & HarvestProcessingDetails("commandline|switches"))
                    End

                End Try

                '--- recursively call the function
                Dim sIncluded As String = BootstrapIncludes(xDoc.OuterXml)
                xDoc.LoadXml(sIncluded)

            Catch ex As System.IO.DirectoryNotFoundException

                '--- file or directory not found
                '--- so just return the same xml
                xDoc.LoadXml(xInclude.OuterXml)

            Catch ex As System.IO.FileNotFoundException

                '--- file or directory not found
                '--- so just return the same xml
                xDoc.LoadXml(xInclude.OuterXml)

            End Try


            '--- replace the include node with the Xml from the include doc
            Dim xTemp As Xml.XmlNode = xJobDefs.ImportNode(xDoc.DocumentElement, True)
            With xInclude.ParentNode
                .AppendChild(xTemp)
                .RemoveChild(xInclude)
            End With

        Next

        '--- return the newly "included" xml
        'xJobDefs.Save("c:\temp\jd.xml")
        Return xJobDefs.OuterXml

    End Function


    ''' <summary>
    ''' Returns a CDATA wrapped string
    ''' </summary>
    ''' <param name="ToBeWrapped">The string to be wrapped up</param>
    ''' <returns></returns>
    ''' <remarks>SDCP 18-Jul-2007</remarks>
    Public Function CDataBlock(ByVal ToBeWrapped As String) As String
        Return "<![CDATA[" & ToBeWrapped & "]]>"
    End Function

    ''' <summary>
    ''' Get current % memory used of total memory available (see "Performance" tab in task manager, Physical Memory)
    ''' </summary>
    ''' <returns>% of total memory used</returns>
    ''' <remarks>Rounded to integer</remarks>
    Function MemoryUsed() As Integer

        Try
            Dim objCS As ManagementObjectSearcher = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
            For Each objMgmt As ManagementObject In objCS.Get
                Dim perfFreeMem As New PerformanceCounter("Memory", "Available Bytes")
                Return CType(Math.Round(100 * perfFreeMem.NextValue() / CType(objMgmt("totalphysicalmemory").ToString(), Double), 0), Integer)
            Next

        Catch ex As Exception
            Dim sMessage As String = "Error in Runner/HelperFunctions/MemoryUsed" & vbCr
            sMessage += ex.Message & vbCr
            sMessage += ex.StackTrace & vbCr
            sMessage += ex.ToString
            MyMessenger.Log(Messenger.LogType.ErrorType, sMessage)
            Return 100

        End Try

    End Function

    ''' <summary>
    ''' Get current CPU Usage, as per task manager, % of total memory used rounded to integer
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CPUUsed() As Integer

        Try

            Dim class1 As ManagementClass = New ManagementClass("Win32_Processor")

            For Each ob As ManagementObject In class1.GetInstances
                Return CType(ob.GetPropertyValue("LoadPercentage").ToString, Integer)
            Next

        Catch ex As Exception
            Return 100
        End Try

    End Function



    ''' <summary>
    ''' escapes to entity references
    ''' </summary>
    ''' <param name="InputString"></param>
    ''' <returns></returns>
    ''' <remarks>SDCP 24-Jan-2007</remarks>
    Public Function Escape(ByVal InputString As String) As String

        Dim x As New Xml.XmlDocument

        Try
            x.LoadXml("<x>" + "</x>")
            x.DocumentElement.InnerText = InputString
            Return x.DocumentElement.InnerXml
        Catch ex As Exception
            Return InputString
        End Try

    End Function


    ''' <summary>
    ''' Gets the named ADS from the file if it exists
    ''' </summary>
    ''' <param name="HostFileFullPath"></param>
    ''' <param name="ADSName"></param>
    ''' <returns>The ADS string, or an error message</returns>
    ''' <remarks>SDCP 30-Jan-2008</remarks>
    Public Function GetADSInfo(ByVal HostFileFullPath As String, ByVal ADSName As String) As String
        Dim fso As Scripting.FileSystemObject = New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim sADSString As String


        '--- check that Hostfile exists
        If Not fso.FileExists(HostFileFullPath) Then
            Err.Raise(vbObjectError + 1001, "HelperFunctions.ReadADS()", "Host file does not exist")
            Return ""
        End If

        '--- check the ADS exists
        If Not fso.FileExists(HostFileFullPath & ":" & ADSName) Then
            Err.Raise(vbObjectError + 1002, "HelperFunctions.ReadADS()", "ADS '" & ADSName & "' does not exist")
            Return ""
        End If

        '--- get the ADS
        ts = fso.OpenTextFile(HostFileFullPath & ":" & ADSName, Scripting.IOMode.ForReading)
        sADSString = ts.ReadAll
        ts.Close()

        '--- return the value 
        Return sADSString

    End Function


    'OBS requires Month/Day/Year, with the hokey month convention
    Public Function ConvertToOBSDate(ByVal MyDate As Date) As String

        Dim sMonths As String
        sMonths = "JAFBMRAPMYJNJLAGSPOCNVDC"
        sMonths = Mid(sMonths, 2 * MyDate.Month - 1, 2)
        Return sMonths & "/" & MyDate.Day.ToString() & "/" & MyDate.Year

    End Function


    Public Function TIBTimeToWindowsTime(ByVal NumericTimestamp As Long) As Date
        Dim lTimestamp As Long
        Dim l1970 As Long

        lTimestamp = NumericTimestamp * 10000
        l1970 = DateAndTime.DateSerial(1970, 1, 1).Ticks
        Return New DateTime(lTimestamp + l1970)

    End Function




    Public Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String, ByVal KeyName As String, _
    ByVal DefaultValue As String) As String
        ' primary version of call gets single value given all parameters
        Dim n As Int32
        Dim sData As String
        sData = Space$(1024) ' allocate some room
        n = GetPrivateProfileString(SectionName, KeyName, DefaultValue, _
        sData, sData.Length, INIPath)
        If n > 0 Then ' return whatever it gave us
            INIRead = sData.Substring(0, n)
        Else
            INIRead = ""
        End If
    End Function

#Region "INIRead Overloads"
    Public Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String, ByVal KeyName As String) As String
        ' overload 1 assumes zero-length default
        Return INIRead(INIPath, SectionName, KeyName, "")
    End Function

    Public Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String) As String
        ' overload 2 returns all keys in a given section of the given file
        Return INIRead(INIPath, SectionName, Nothing, "")
    End Function

    Public Function INIRead(ByVal INIPath As String) As String
        ' overload 3 returns all section names given just path
        Return INIRead(INIPath, Nothing, Nothing, "")
    End Function
#End Region

    Public Sub INIWrite(ByVal INIPath As String, ByVal SectionName As String, _
    ByVal KeyName As String, ByVal TheValue As String)
        Call WritePrivateProfileString(SectionName, KeyName, TheValue, INIPath)
    End Sub

    Public Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String, _
    ByVal KeyName As String) ' delete single line from section
        Call WritePrivateProfileString(SectionName, KeyName, Nothing, INIPath)
    End Sub

    Public Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String)
        ' delete section from INI file
        Call WritePrivateProfileString(SectionName, Nothing, Nothing, INIPath)
    End Sub


End Module
