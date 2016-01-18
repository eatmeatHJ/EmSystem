
''' <summary>
''' An example configuration section class.
''' </summary>
Public Class EmDBSection
    Inherits ConfigurationSection

#Region "Static Fields"
    Private Shared IsCompany As ConfigurationProperty
    Private Shared IsForm As ConfigurationProperty
    Private Shared IsSaveSql As ConfigurationProperty
    Private Shared IsSqlAtk As ConfigurationProperty
    Private Shared ConnStr As ConfigurationProperty

    Private Shared s_properties As ConfigurationPropertyCollection
#End Region

#Region "Constructors"
    Shared Sub New()
        ' Predefine properties here
        IsCompany = New ConfigurationProperty("IsCompany", GetType(Boolean), Nothing, ConfigurationPropertyOptions.None)
        IsForm = New ConfigurationProperty("IsForm", GetType(Boolean), False, ConfigurationPropertyOptions.None)
        IsSaveSql = New ConfigurationProperty("IsSaveSql", GetType(Boolean), Nothing, ConfigurationPropertyOptions.None)
        IsSqlAtk = New ConfigurationProperty("IsSqlAtk", GetType(Boolean), Nothing, ConfigurationPropertyOptions.None)
        ConnStr = New ConfigurationProperty("ConnStr", GetType(String), Nothing, ConfigurationPropertyOptions.None)
        s_properties = New ConfigurationPropertyCollection()

        s_properties.Add(IsCompany)
        s_properties.Add(IsForm)
        s_properties.Add(IsSaveSql)
        s_properties.Add(IsSqlAtk)
        s_properties.Add(ConnStr)
    End Sub
#End Region

#Region "Properties"
    <ConfigurationProperty("IsCompany")>
    Public ReadOnly Property GetIsCompany() As Boolean
        Get
            Return Item(IsCompany)
        End Get
    End Property

    <ConfigurationProperty("IsForm")>
    Public ReadOnly Property GetIsForm() As Boolean
        Get
            Return Item(IsForm)
        End Get
    End Property

    <ConfigurationProperty("IsSaveSql")>
    Public ReadOnly Property GetIsSaveSql() As Boolean
        Get
            Return Item(IsSaveSql)
        End Get
    End Property

    <ConfigurationProperty("IsSqlAtk")>
    Public ReadOnly Property GetIsSqlAtk() As Boolean
        Get
            Return Item(IsSqlAtk)
        End Get
    End Property

    <ConfigurationProperty("ConnStr")>
    Public ReadOnly Property GetConnStr() As String
        Get
            Return Item(ConnStr)
        End Get
    End Property

    Protected Overrides ReadOnly Property Properties() As ConfigurationPropertyCollection
        Get
            Return s_properties
        End Get
    End Property
#End Region
End Class
