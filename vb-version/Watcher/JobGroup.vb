Public Class JobGroup
    Inherits List(Of Job)

    '--- private variables
    Private _name As String


    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

End Class
