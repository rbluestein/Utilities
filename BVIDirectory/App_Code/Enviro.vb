Public Class Enviro
#Region " Declarations "
    Private cAppTimedOut As Boolean
    Private cLoggedInUserID As String
    Private cLoggedInUserName As String
    Private cComputerId As String
    Private cLoginLocationID As String
    Private cLoginRoleCatgy As RoleCatgyEnum
    Private cDBHost As String
    Private cTestProd As TestProdEnum
    Private cConnectionStringTemplate As String
    Private cInit As Boolean
    Private cLastPageLoad As DateTime
    Private cApplicationPath As String
    Private cLoginIP As String
    Private cLogFileFullPath As String
#End Region

#Region " Constants "
    Private Const cVersionNumber As String = "1.0"
    Private Const cDefaultDatabase As String = "UserManagement"
    Private Const cDBTimeout As Integer = 10
    Private Const cAppTimeout As Integer = 360000
    Private Const cMaxDisplayRecords As Integer = 500
    Private Const cExcessiveRecordAmount As Integer = 1000
    Private Const cRecordMaximum As Integer = 5000
    Private Const cLogRetentionDays = 30
#End Region


    Public ReadOnly Property VersionNumber() As String
        Get
            Return cVersionNumber
        End Get
    End Property
    Public ReadOnly Property MaxDisplayRecords() As Integer
        Get
            Return cMaxDisplayRecords
        End Get
    End Property
    Public ReadOnly Property ExcessiveRecordAmount() As Integer
        Get
            Return cExcessiveRecordAmount
        End Get
    End Property
    Public ReadOnly Property RecordMaximum() As Integer
        Get
            Return cRecordMaximum
        End Get
    End Property
    Public ReadOnly Property LogRetentionDays() As Integer
        Get
            Return cLogRetentionDays
        End Get
    End Property

    Public Property Init() As Boolean
        Get
            Return cInit
        End Get
        Set(ByVal Value As Boolean)
            cInit = Value
        End Set
    End Property
    Public Property LastPageLoad() As DateTime
        Get
            Return cLastPageLoad
        End Get
        Set(ByVal Value As DateTime)
            cLastPageLoad = Value
        End Set
    End Property
    Public Property LoggedInUserID() As String
        Get
            If cLoggedInUserID = Nothing Then
                Return GetAlternateLoggedInUserID()
            Else
                Return cLoggedInUserID
            End If
        End Get
        Set(ByVal Value As String)
            cLoggedInUserID = Value
        End Set
    End Property

    Public Property LoggedInUserName() As String
        Get
            Return cLoggedInUserName
        End Get
        Set(ByVal Value As String)
            cLoggedInUserName = Value
        End Set
    End Property

    Private Function GetAlternateLoggedInUserID() As String
        Dim UserID As String
        UserID = HttpContext.Current.User.Identity.Name.ToString
        UserID = UserID.Substring(InStr(UserID, "\", CompareMethod.Binary))
        Return UserID
    End Function

    Public Property LoginIP() As String
        Get
            Return cLoginIP
        End Get
        Set(ByVal Value As String)
            cLoginIP = Value
        End Set
    End Property

    Public Property LoginLocationID() As String
        Get
            Return cLoginLocationID
        End Get
        Set(ByVal Value As String)
            cLoginLocationID = Value
        End Set
    End Property

    Public Property LogInRoleCatgy() As RoleCatgyEnum
        Get
            Return cLoginRoleCatgy
        End Get
        Set(ByVal Value As RoleCatgyEnum)
            cLoginRoleCatgy = Value
        End Set
    End Property

    Public Property ApplicationPath() As String
        Get
            Return cApplicationPath
        End Get
        Set(ByVal Value As String)
            cApplicationPath = Value
        End Set
    End Property

    Public Property LogFileFullPath() As String
        Get
            Return cLogFileFullPath
        End Get
        Set(ByVal Value As String)
            cLogFileFullPath = Value
        End Set
    End Property


    Public Property DBHost() As String
        Get
            Return cDBHost
        End Get
        Set(ByVal Value As String)
            cDBHost = Value
        End Set
    End Property

    Public Property TestProd() As TestProdEnum
        Get
            Return cTestProd
        End Get
        Set(ByVal Value As TestProdEnum)
            cTestProd = Value
        End Set
    End Property

    Public ReadOnly Property AppTimeout() As Integer
        Get
            Return cAppTimeout
        End Get
    End Property

    Public Property AppTimedOut() As Boolean
        Get
            Return cAppTimedOut
        End Get
        Set(ByVal Value As Boolean)
            cAppTimedOut = Value
        End Set
    End Property

    Public ReadOnly Property DefaultDatabase() As String
        Get
            Return cDefaultDatabase
        End Get
    End Property

    Public Property ConnectionStringTemplate() As String
        Get
            Return cConnectionStringTemplate
        End Get
        Set(ByVal Value As String)
            cConnectionStringTemplate = Value
        End Set
    End Property

#Region " ConnectionString "
    Public Function GetConnectionString(ByVal DBHost As String, ByVal Database As String) As String
        Return Replace(cConnectionStringTemplate, "|", Database) & DBHost
    End Function

    Public Function GetConnectionString() As String
        Return Replace(cConnectionStringTemplate, "|", cDefaultDatabase) & cDBHost
    End Function
#End Region

    Public Sub MakeCookie(ByVal Page As Page)
        Dim SessionCookie As HttpCookie
        SessionCookie = New HttpCookie("BVICCM", cLoggedInUserID)
        Page.Response.Cookies.Add(SessionCookie)
    End Sub

    Public Sub AuthenticateRequest_Timeout(ByRef Page As System.Web.UI.Page)
        Dim ts As TimeSpan
        Dim EnviroLoadedInd As Boolean
        Dim CookieInd As Boolean
        Dim SessionInd As Boolean
        Dim WrongPageEntry_ServerTimeoutErrorInd As Boolean
        Dim SessionTimeoutErrorInd As Boolean
        Dim CookieErrorInd As Boolean

        'EnviroLoaded         CookiePresent        SessionActive       Action
        'Yes                      Yes                       Yes                       Normal operation
        'Yes                      Yes                       No                        Session timeout
        'Yes                      No                        Yes                       No cookie
        'Yes                      No                        No                        Session timeout
        'No                       Yes                       Yes                       Server timeout
        'No                       Yes                       No                        Server timeout
        'No                       No                        Yes                       Server timeout
        'No                       No                        No                        Wrong page entry / Server timeout

        If cConnectionStringTemplate = Nothing Then
            EnviroLoadedInd = False
        Else
            EnviroLoadedInd = True
        End If
        If IsNothing(Page.Request.Cookies.Item("BVICCM")) Then
            CookieInd = False
        Else
            CookieInd = True
        End If
        If cAppTimedOut Then
            SessionInd = False
        Else
            SessionInd = True
        End If

        ts = Date.Now.Subtract(cLastPageLoad)
        If ts.TotalSeconds > cAppTimeout Then
            SessionInd = False
        Else
            SessionInd = True
            cLastPageLoad = Date.Now
        End If

        If EnviroLoadedInd Then
            If CookieInd Then
                If SessionInd Then
                    ' good to go
                Else
                    SessionTimeoutErrorInd = True
                End If
            Else
                If SessionInd Then
                    CookieErrorInd = True
                Else
                    SessionTimeoutErrorInd = True
                End If
            End If
        Else
            WrongPageEntry_ServerTimeoutErrorInd = True
        End If

        If WrongPageEntry_ServerTimeoutErrorInd Then
            Throw New Exception("Either an incorrect page entry or a server time out has occurred. Please close the application and log back into http://netserver.benefitvision.com/BVIDirectoryr/.")
        ElseIf SessionTimeoutErrorInd Then
            Throw New Exception("Application session has timed out. Please close the application and log back in.  Last page load: " & cLastPageLoad.ToString("MM/dd/yyyy hh:mm tt") & ". This page load: " & Date.Now.ToString("MM/dd/yyyy hh:mm tt") & ". Total minutes: " & CInt(ts.TotalMinutes) & ".")
        ElseIf CookieErrorInd Then
            Throw New Exception("No cookie.")
        End If

    End Sub

    'Public Sub ORIGAuthenticateRequest_Timeout(ByRef Page As System.Web.UI.Page)
    '    Try

    '        ' ___ Validate cookie
    '        If IsNothing(Page.Request.Cookies.Item("BVICCM")) Then
    '            'Throw New Exception("No Cookie Present.")
    '            Throw New Exception("No cookie present. The most likely cause is that you are launching the application from the wrong page.")
    '        Else
    '            If Page.Request.Cookies.Item("BVICCM").Value <> cLoggedInUserID Then
    '                Throw New Exception("No cookie present. The most likely cause is that you are launching the application from the wrong page.")
    '            End If
    '        End If

    '        ' ___ Check for time out
    '        ValidateAppTimeout()

    '    Catch ex As Exception
    '        Throw New Exception("Error #820: Enviro ValidateRequest. " & ex.Message, ex)
    '    End Try
    'End Sub

    'Private Sub ValidateAppTimeout()
    '    Dim ts As TimeSpan

    '    Try
    '        If cAppTimedOut Then
    '            Throw New Exception("posttimeout")
    '        End If

    '        ts = Date.Now.Subtract(cLastPageLoad)
    '        If ts.TotalSeconds > cAppTimeout Then
    '            Throw New Exception("Application Timeout.")
    '            'Throw New Exception("Application has timed out. Please close the application and log back in.  Last page load: " & cLastPageLoad.ToString("MM/dd/yyyy hh:mm tt") & ". This page load: " & Date.Now.ToString("MM/dd/yyyy hh:mm tt") & ". Total minutes: " & CInt(ts.TotalMinutes) & ".")
    '        Else
    '            cLastPageLoad = Date.Now
    '        End If

    '    Catch ex As Exception
    '        Throw New Exception("Error #821: Enviro ValidateAppTimeout. " & ex.Message, ex)
    '    End Try
    'End Sub
End Class