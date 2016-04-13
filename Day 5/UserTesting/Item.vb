Public Class Item
    Private _name As String
    Private _price As Double
    Private _qty As Integer
    Private _desc As String

    Public Property Name As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Price As Double
        Get
            Return _price
        End Get
        Set(ByVal value As Double)
            _price = value
        End Set
    End Property

    Public Property Qty As Integer
        Get
            Return _qty
        End Get
        Set(ByVal value As Integer)
            _qty = value
        End Set
    End Property

    Public Property Desc As String
        Get
            Return _desc
        End Get
        Set(ByVal value As String)
            _desc = value
        End Set
    End Property

End Class
