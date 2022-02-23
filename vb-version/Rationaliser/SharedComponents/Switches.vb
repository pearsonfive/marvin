Public Class Switches

    '--- private variables
    Friend _switchesHashTable As Hashtable


    ''' <summary>
    ''' static constructor to create the switches class
    ''' </summary>
    ''' <remarks>SDCP 20-Oct-2006</remarks>
    Sub New()
        Dim switchName As String
        Dim switchValue As String

        _switchesHashTable = New Hashtable

        For Each switch As String In My.Application.CommandLineArgs

            switchName = ""
            switchValue = ""

            Try

                switchName = switch.Split("="c)(0).Substring(1).ToLower
                switchValue = switch.Split("="c)(1)

                'By having this in the Try/Catch it handles duplicate ht entries (by ignoring them)
                _switchesHashTable.Add(switchName, switchValue)

            Catch ex As Exception
            End Try

        Next

        If Not _switchesHashTable.ContainsKey("environment") Then
            _switchesHashTable.Add("environment", "development")
        End If

    End Sub


    ''' <summary>
    ''' default property that allows extraction of 1 value
    ''' </summary>
    ''' <param name="SwitchName">The name of the switch to look for</param>
    ''' <value></value>
    ''' <returns>the value, or empty string if the switch does not exist</returns>
    ''' <remarks>SDCP 20-Oct-2006</remarks>
    Default Public Property Item(ByVal SwitchName As String) As String
        Get
            Dim oValue As Object = Nothing
            Try
                oValue = _switchesHashTable(SwitchName)
            Catch ex As Exception
            End Try

            '--- return the value
            If oValue IsNot Nothing Then
                Return oValue.ToString
            Else
                Return ""
            End If

        End Get
        Set(ByVal value As String)

            Try
                _switchesHashTable(SwitchName) = value
            Catch ex As Exception

            End Try

        End Set
    End Property

    Public ReadOnly Property FullCommandLine() As String
        Get
            Return Microsoft.VisualBasic.Command()
        End Get
    End Property

End Class
