Public Class PageSession
    Private cSortReference As String = String.Empty
    Private cFilterOnOffState As String = String.Empty
    Private cPageInitiallyLoaded As Boolean
    Private cPageReturnOnLoadMessasge As String = String.Empty
    Private cSql As String
    Private cExcessiveRecordsWarningInEffect As Boolean
    Private cInitialReportDataSuppressInEffect As Boolean
    Private cSessionID As String
    Private cUserConfirmationResolvedAsTrue As Boolean

    Public Property SortReference() As String
        Get
            Return cSortReference
        End Get
        Set(ByVal Value As String)
            cSortReference = Value
        End Set
    End Property
    Public Property PageReturnOnLoadMessage() As String
        Get
            Return cPageReturnOnLoadMessasge
        End Get
        Set(ByVal Value As String)
            cPageReturnOnLoadMessasge = Value
        End Set
    End Property
    Public Property FilterOnOffState() As String
        Get
            Return cFilterOnOffState
        End Get
        Set(ByVal Value As String)
            cFilterOnOffState = Value
        End Set
    End Property
    Public Property PageInitiallyLoaded() As Boolean
        Get
            Return cPageInitiallyLoaded
        End Get
        Set(ByVal Value As Boolean)
            cPageInitiallyLoaded = Value
        End Set
    End Property

    Public Property ExcessiveRecordsWarningInEffect() As Boolean
        Get
            Return cExcessiveRecordsWarningInEffect
        End Get
        Set(ByVal Value As Boolean)
            cExcessiveRecordsWarningInEffect = Value
        End Set
    End Property
    Public Property Sql() As String
        Get
            Return cSql
        End Get
        Set(ByVal Value As String)
            cSql = Value
        End Set
    End Property
    Public Property InitialReportDataSuppressInEffect() As Boolean
        Get
            Return cInitialReportDataSuppressInEffect
        End Get
        Set(ByVal Value As Boolean)
            cInitialReportDataSuppressInEffect = Value
        End Set
    End Property
End Class

Public Class IndexSession
    Inherits PageSession

    'Private cUserID As String = String.Empty
    Private cUserIDFilter As String = String.Empty
    Private cFullNameFilter As String = String.Empty
    Private cStatusCodeFilter As String = String.Empty
    Private cRoleFilter As String = String.Empty
    Private cCompanyIDFilter As String = String.Empty
    Private cLocationIDFilter As String = String.Empty
    Private cJumpToEnrollerID As String = String.Empty

    Public Sub New()
        MyBase.new()
        'cSessionID = Guid.NewGuid.ToString
    End Sub

    Public Property JumpToEnrollerID() As String
        Get
            Return cJumpToEnrollerID
        End Get
        Set(ByVal Value As String)
            cJumpToEnrollerID = Value
        End Set
    End Property

    'Public Property UserID() As String
    '    Get
    '        Return cUserID
    '    End Get
    '    Set(ByVal Value As String)
    '        cUserID = Value
    '    End Set
    'End Property
    Public Property UserIDFilter() As String
        Get
            Return cUserIDFilter
        End Get
        Set(ByVal Value As String)
            cUserIDFilter = Value
        End Set
    End Property
    Public Property FullNameFilter() As String
        Get
            Return cFullNameFilter
        End Get
        Set(ByVal Value As String)
            cFullNameFilter = Value
        End Set
    End Property
    Public Property StatusCodeFilter() As String
        Get
            Return cStatusCodeFilter
        End Get
        Set(ByVal Value As String)
            cStatusCodeFilter = Value
        End Set
    End Property
    Public Property RoleFilter() As String
        Get
            Return cRoleFilter
        End Get
        Set(ByVal Value As String)
            cRoleFilter = Value
        End Set
    End Property
    Public Property CompanyIDFilter() As String
        Get
            Return cCompanyIDFilter
        End Get
        Set(ByVal Value As String)
            cCompanyIDFilter = Value
        End Set
    End Property
    Public Property LocationIDFilter() As String
        Get
            Return cLocationIDFilter
        End Get
        Set(ByVal Value As String)
            cLocationIDFilter = Value
        End Set
    End Property
End Class