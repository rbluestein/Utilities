Imports System.Data

Partial Class Index
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Private cIndexSess As IndexSession
#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim PageMode As PageMode
        Dim Action As String
        Dim DG As DG

        Try
            Try

                ' ___ Get Enviro from Session
                cEnviro = Session("Enviro")

                ' ___ Instantiate Common
                cCommon = New Common

                ' ___ Get the enroller session object
                cIndexSess = Session("IndexSession")

                ' ___ Load the application settings
                If Not cEnviro.Init Then
                    LoadEnviro()
                Else
                    cEnviro.AuthenticateRequest_Timeout(Me)
                End If

                ' ___ Right check
                cRights = New RightsClass(cEnviro, Page)
                Dim RightsRqd(0) As String
                RightsRqd.SetValue(RightsClass.DirectoryView, 0)
                cRights.HasSufficientRights(RightsRqd, True, Page)
                'lblCurrentRights.Text = cCommon.GetCurRightsAndTopicsHidden(cRights.RightsColl)

                ' ___ Get the page mode 
                PageMode = cCommon.GetPageMode(Page, cIndexSess)

                ' ___ Load the page session variables
                LoadVariables(PageMode)

                ' ___ Initialize the datagrid
                DG = DefineDataGrid()

                ' ___ Execute action
                Select Case PageMode
                    Case PageMode.Initial
                        DisplayPage(PageMode, DG, DG.OrderByType.Initial)

                    Case PageMode.Postback
                        Action = Request.Form("hdAction")
                        Select Case Action
                            Case "Sort"
                                DisplayPage(PageMode, DG, DG.OrderByType.Field, Request.Form("hdSortField"))

                            Case "ApplyFilter"
                                DisplayPage(PageMode, DG, DG.OrderByType.Recurring)
                        End Select

                    Case PageMode.ReturnFromChild, PageMode.CalledByOther
                        DisplayPage(PageMode, DG, DG.OrderByType.ReturnToPage)
                        If cIndexSess.PageReturnOnLoadMessage <> Nothing Then
                            litMsg.Text = "<script language='javascript'>alert('" & cIndexSess.PageReturnOnLoadMessage & "')</script>"
                            cIndexSess.PageReturnOnLoadMessage = Nothing
                        End If
                End Select

                ' ___ Display enviroment
                'PageCaption.Text = cCommon.GetPageCaption
                'litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'>"
                If cIndexSess.JumpToEnrollerID <> String.Empty Then
                    litJumpToAnchor.Text = "window.location.hash='#" & cIndexSess.JumpToEnrollerID & "'"
                End If

                litDBHost.Text = "Site: " & cEnviro.DBHost

            Catch ex As Exception
                Throw New Exception("Error #102: Index Page_Load. " & ex.Message, ex)
            End Try
        Catch ex As Exception
            Dim ErrorObj As New ErrorObj(ex)
        End Try
    End Sub

    Private Sub LoadVariables(ByVal PageMode As PageMode)
        Try

            Select Case PageMode
                Case PageMode.Initial
                    ' No action
                Case PageMode.CalledByOther, PageMode.ReturnFromChild
                    ' No action
                Case PageMode.Postback
                    ' ___ Update session variables with those that the user may have changed
                    'cIndexSess.UserID = Replace(Request.Form("hdUserID"), "~", "'")
                    'cIndexSess.UserIDFilter = Request.Form("txtUserID")
                    'cIndexSess.FullNameFilter = Request.Form("txtFullName")
                    cIndexSess.RoleFilter = Request.Form("ddRole")
                    cIndexSess.JumpToEnrollerID = Request.Form("hdEnrollerID")
                    cIndexSess.LocationIDFilter = Request.Form("ddLocationID")
            End Select
        Catch ex As Exception
            Throw New Exception("Error #104: Index LoadVariables. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub LoadEnviro()
        Dim LoggedInUserID As String
        Dim HTTPHost As String
        Dim dt As DataTable
        Dim Querypack As DBase.QueryPack

        Try

            cEnviro.Init = True
            cEnviro.LastPageLoad = Date.Now

            ' ___ LoggedInUserIDth
            If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
                'LoggedInUserID = "rbluestein"
                'LoggedInUserID = "lbrogan"
                'LoggedInUserID = "rsmith"
                'LoggedInUserID = "jwatroba"
                'LoggedInUserID = "eenroller"
                ' LoggedInUserID = "bgriffin"
                ' LoggedInUserID = "eallbritton"
                ' LoggedInUserID = "jkleiman"
                ' LoggedInUserID = "mtarpley"
                ' LoggedInUserID = "pmauldin"
                'LoggedInUserID = "pblack"
                'LoggedInUserID = "ssupervisor"
                'LoggedInUserID = "bwatroba"
                'LoggedInUserID = "cwright"
                'LoggedInUserID = "jpenny"
                'LoggedInUserID = "eenroller"
                'LoggedInUserID = "flinden"
                'LoggedInUserID = "sboothe"
                LoggedInUserID = "memby"
                'LoggedInUserID = "hbruce"
                'LoggedInUserID = "pmauldin"
                ' LoggedInUserID = "dwalter"
                'LoggedInUserID = "dtopper"
                'LoggedInUserID = "kcarter"
                'LoggedInUserID = "flinden"
                'LoggedInUserID = "awickenheiser"


            Else
                LoggedInUserID = HttpContext.Current.User.Identity.Name.ToString
                LoggedInUserID = LoggedInUserID.Substring(InStr(LoggedInUserID, "\", CompareMethod.Binary))
            End If
            cEnviro.LoggedInUserID = LoggedInUserID

            ' ___ DBHost
            cEnviro.DBHost = ConfigurationManager.AppSettings("DBHost")

            ' ___ Environment
            If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
                cEnviro.TestProd = TestProdEnum.Production
            Else
                HTTPHost = Page.Request.ServerVariables("http_host").ToLower
                If HTTPHost = "netserver.benefitvision.com" Then
                    cEnviro.TestProd = TestProdEnum.Production
                Else
                    cEnviro.TestProd = TestProdEnum.Production
                End If
            End If

            cEnviro.ConnectionStringTemplate = "User ID=BVI_SQL_SERVER;Password=noisivtifeneb;database=|;server="
            cEnviro.LoginIP = Page.Request.ServerVariables("REMOTE_ADDR")
            cEnviro.ApplicationPath = Page.Server.MapPath(Page.Request.ApplicationPath)
            cEnviro.LogFileFullPath = Page.Server.MapPath(Page.Request.ApplicationPath) & "\BVIDirectory.txt"

            Querypack = cCommon.GetDTWithQueryPack("SELECT Count (*) FROM UserManagement..Users")
            If Not Querypack.Success Then
                Throw New Exception("Error #115a: Index.LoadEnviro. UnableToConnect")
            End If

            dt = cCommon.GetDT("SELECT LastName + ', ' + FirstName FROM UserManagement..Users WHERE UserID = '" & LoggedInUserID & "'")
            If Not IsDBNull(dt.Rows(0)(0)) Then
                cEnviro.LoggedInUserName = dt.Rows(0)(0)
            End If

            ' __ Login LocationID and LoginRole
            dt = cCommon.GetDT("SELECT Role, LocationID FROM Users WHERE UserID ='" & cEnviro.LoggedInUserID & "'", cEnviro.DBHost, "UserManagement")
            cEnviro.LoginLocationID = dt.Rows(0)("LocationID")

            Select Case dt.Rows(0)("Role").ToUpper
                Case "IT", "ADMIN", "ADMIN LIC"
                    cEnviro.LogInRoleCatgy = RoleCatgyEnum.Other
                Case "ENROLLER"
                    cEnviro.LogInRoleCatgy = RoleCatgyEnum.Enroller
                Case "SUPERVISOR"
                    cEnviro.LogInRoleCatgy = RoleCatgyEnum.Supervisor
            End Select

            cEnviro.MakeCookie(Me)

        Catch ex As Exception
            Throw New Exception("Error #115b: Index.LoadEnviro. " & ex.Message, ex)
        End Try
    End Sub


    Private Function DefineDataGrid() As DG
        Try

            Dim DG As New DG("UserID", cCommon, cRights, True, "EmbeddedTableDef", "FullName", DG.DefaultSortDirectionEnum.Ascending)

            DG.AddAnchorObject(Nothing, Nothing, Nothing)
            DG.AddDataBoundColumn("UserID", "UserID", "Enroller ID", "UserID", False, Nothing, Nothing, "align='left'")
            'DG.AddDataBoundColumn("FullName", "FullName", "Name", "LastName+FirstName+MI", True, Nothing, Nothing, "align='left' width='100px'")
            DG.AddDataBoundColumn("FullName", "FullName", "Name", "FullName", True, Nothing, Nothing, "align='left' width='120px'")
            DG.AddDataBoundColumn("PhoneExtension", "PhoneExtension", "Ext", "PhoneExtension", True, Nothing, Nothing, "align='left' width='50px'")
            DG.AddDataBoundColumn("Role", "Role", "Role", "Role", True, Nothing, Nothing, "align='left' width='50px'")
            DG.AddDataBoundColumn("LocationID", "LocationID", "Location", "LocationID", True, Nothing, Nothing, "align='left' width='50px'")

            ' ___ Build the filter
            Dim Filter As DG.Filter
            Filter = DG.AttachFilter(DG.FilterOperationMode.FilterAlwaysOn, DG.FilterInitialShowHideEnum.FilterInitialShow, DG.RecordsInitialShowHideEnum.RecordsInitialShow)
            'Filter.AddNameTextbox("FullName", "FullName", "LastName", "FirstName", 48, Nothing)
            Filter.AddDropdown("Role", "Role")
            Filter.AddDropdown("LocationID", "LocationID")

            Return DG

        Catch ex As Exception
            Throw New Exception("Error #106: Index DefineDataGrid. " & ex.Message, ex)
        End Try
    End Function

    Private Sub DisplayPage(ByVal PageMode As PageMode, ByVal DG As DG, ByVal OrderByType As DG.OrderByType, Optional ByVal OrderByField As String = Nothing)
        Dim ViewOrDownload As String = "View"

        Try

            ' ___ Handle the filter
            HandleFilter(DG, PageMode)

            ' ___ Handle the sort
            If cIndexSess.SortReference <> Nothing Then
                DG.UpdateSortReference(cIndexSess.SortReference)
            End If
            DG.SetSortElements(OrderByField, OrderByType)

            ' ___ Handle the data
            HandleData(DG, PageMode, OrderByType, ViewOrDownload)

            ' ___ Set the FilterOnOffState
            cIndexSess.FilterOnOffState = "on"

            ' ___ Set the last field sorted and sort direction in the sort reference
            cIndexSess.SortReference = DG.GetSortReference

        Catch ex As Exception
            Throw New Exception("Error #110: Index DisplayPage. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleFilter(ByRef DG As DG, ByVal PageMode As PageMode)
        Dim i As Integer
        Dim dt As DataTable

        Try

            ' ___ Get a filter reference
            Dim Filter As DG.Filter
            Filter = DG.GetFilter

            ' ___ Location
            Filter("LocationID").AddDropdownItem("", "", True)
            Filter("LocationID").AddDropdownItem("HBG", "HBG", False)
            Filter("LocationID").AddDropdownItem("LA", "LA", False)
            Filter("LocationID").AddDropdownItem("REM", "REM", False)

            If PageMode <> PageMode.Initial Then
                Filter.Coll("LocationID").SetFilterValue(cIndexSess.LocationIDFilter)
            End If

            ' ___ Role
            Filter("Role").AddDropdownItem("", "", True)
            Filter("Role").AddDropdownItem("Admin", "Admin", False)
            Filter("Role").AddDropdownItem("Admin Lic", "Admin Lic", False)
            Filter("Role").AddDropdownItem("Enroller", "Enroller", False)
            Filter("Role").AddDropdownItem("IT", "IT", False)
            Filter("Role").AddDropdownItem("Supervisor", "Supervisor", False)

            If PageMode <> PageMode.Initial Then
                Filter.Coll("Role").SetFilterValue(cIndexSess.RoleFilter)
            End If

        Catch ex As Exception
            Throw New Exception("Error #111: Index HandleFilter. " & ex.Message, ex)
        End Try
    End Sub

    Private Sub HandleData(ByRef DG As DG, ByVal PageMode As PageMode, ByVal OrderByType As DG.OrderByType, ByVal ViewOrDownload As String)
        Dim dt As DataTable
        Dim RecordCount As Integer
        Dim Coll As Collection
        Dim SuppressDisplayData As Boolean
        Dim EmbeddedMessage As String = Nothing
        Dim sb As New System.Text.StringBuilder
        Dim DownloadPathColl As Collection
        Dim HTMLReport As Boolean
        Dim DownloadReport As Boolean
        Dim ExceedsRecordMaximum As Boolean
        Dim IgnoreExcessiveRecords As Boolean
        Dim NewQuery As Boolean
        Dim PerformNextTest As Boolean

        Try

            ' ___ #1: FIGURE OUT WHAT'S GOING ON

            ' ___ Get the recordcount
            Coll = GetQueryInfo("RecordCount", DG, OrderByType)
            RecordCount = Coll("RecordCount")


            ' ___ Test #1: InitialReportDataSuppress
            ' User opens this page. The datagrid suppresses the initial display of the data. 
            ' If the user navigates to a different page and then returns to this page, PageMode is no longer set
            ' to initial and the initial data is permitted to display. The Sql attempts to return all of the records the sql statement
            ' with no restrictions. A postback is required to enable the display of the data.

            If PageMode = PageMode.Initial AndAlso DG.RecordsInitialShowHide = DG.RecordsInitialShowHideEnum.RecordsInitialHide Then
                cIndexSess.InitialReportDataSuppressInEffect = True
            ElseIf PageMode = PageMode.Postback AndAlso cIndexSess.InitialReportDataSuppressInEffect Then
                cIndexSess.InitialReportDataSuppressInEffect = False
            End If
            If Not cIndexSess.InitialReportDataSuppressInEffect Then
                PerformNextTest = True
            End If


            ' ___ Test #2: Exceeds record maximum test
            If PerformNextTest Then
                If RecordCount > cEnviro.RecordMaximum Then
                    ExceedsRecordMaximum = True
                    PerformNextTest = False
                End If
            End If

            ' ___ Test #3: Ignore excessive records warning
            ' If an excessive records warning was put into effect and the user has chosen to proceed with the same query, ignore and reset the warning.
            If PerformNextTest Then
                If cIndexSess.ExcessiveRecordsWarningInEffect Then
                    Coll = GetQueryInfo("Sql", DG, OrderByType)
                    If StrComp(Coll("Sql"), cIndexSess.Sql, CompareMethod.Text) = 0 Then
                        IgnoreExcessiveRecords = True
                        PerformNextTest = False
                    Else
                        NewQuery = True
                    End If
                Else
                    NewQuery = True
                End If

                ' ___ Reset the excessive record warning properties
                If cIndexSess.ExcessiveRecordsWarningInEffect Then
                    cIndexSess.ExcessiveRecordsWarningInEffect = False
                    cIndexSess.Sql = String.Empty
                End If
            End If

            ' ___ Test #4: Excessive records test
            If PerformNextTest Then
                If RecordCount > cEnviro.ExcessiveRecordAmount Then
                    Coll = GetQueryInfo("Sql", DG, OrderByType)
                    cIndexSess.ExcessiveRecordsWarningInEffect = True
                    cIndexSess.Sql = Coll("Sql")
                End If
            End If

            ' ___ DIRECT THE REPORT
            If ViewOrDownload = "Download" Then
                DownloadReport = True
            Else
                HTMLReport = True
            End If
            If cIndexSess.InitialReportDataSuppressInEffect Then
                HTMLReport = True
                DownloadReport = False
                SuppressDisplayData = True
            ElseIf ExceedsRecordMaximum Then
                HTMLReport = True
                DownloadReport = False
                SuppressDisplayData = True
                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">&nbsp;&nbsp;Report contains " & RecordCount & " records. Respecify report.</td>"  '"<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">" & EmbeddedMessage & "</td>"
            ElseIf IgnoreExcessiveRecords Then
                HTMLReport = False
                DownloadReport = True
            ElseIf cIndexSess.ExcessiveRecordsWarningInEffect Then
                HTMLReport = True
                DownloadReport = False
                SuppressDisplayData = True
                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">&nbsp;&nbsp;Report contains " & RecordCount & " records. Proceed or respecify report.</td>"  '"<td style=""font: 10pt Arial, Helvetica, sans-serif;color:red"">" & EmbeddedMessage & "</td>"
            End If
            If DownloadReport Then
                SuppressDisplayData = True
                EmbeddedMessage = "<td style=""font: 10pt Arial, Helvetica, sans-serif"">&nbsp;&nbsp;Click here to <a href=""" & DownloadPathColl("RelPath") & """>download</a> your CSV file.</td>"
            End If


            ' ___ EXECUTE THE REPORT

            ' ___ Get the data
            If (HTMLReport And (Not SuppressDisplayData)) OrElse DownloadReport Then
                Coll = GetQueryInfo("Data", DG, OrderByType)
                dt = Coll("Data")
            End If

            ' ___ Process the download
            If DownloadReport Then
                cCommon.PrintCSVVersionLocal(dt, DownloadPathColl("AbsPath"), Nothing)
            End If

            ' ___ Process the html
            If SuppressDisplayData Then
                dt = Nothing
            End If
            litDG.Text = DG.GetText(dt, Request, EmbeddedMessage)

        Catch ex As Exception
            Throw New Exception("Error #112: Index DisplayPage. " & ex.Message, ex)
        End Try
    End Sub

    Private Function GetQueryInfo(ByVal InfoType As String, ByVal DG As DG, ByVal OrderByType As DG.OrderByType) As Collection
        Dim sb As New System.Text.StringBuilder
        Dim Sql As String
        Dim dt As DataTable
        Dim ShowFilter As Boolean
        Dim Coll As New Collection
        Dim Pos As Integer
        Dim NonFilterWhereClause As String
        Dim AccessLevel As RightsClass.AccessLevelEnum
        Dim Role As String
        Dim LocationID As String
        Dim QueryPack As DBase.QueryPack

        Try


            ' ___ This method is requested to return one of three items:
            ' (1) a recordcount, (2) a datatable, or (3) a sql statement.

            ' ___ Get the security settings for the logged in user
            cRights.GetSecurityFlds(AccessLevel, cEnviro.LogInRoleCatgy)

            If InfoType = "RecordCount" Then
                sb.Append("SELECT Count (*) ")

            ElseIf InfoType = "Data" Or InfoType = "Sql" Then
                sb.Append("SELECT UserId, ")
                'sb.Append("LastName, FirstName, ")
                sb.Append("LTrim(RTrim(LastName)) + ', ' + LTrim(RTrim(FirstName))  FullName,  ")
                'sb.Append("LTrim(RTrim(FirstName)) + ' ' + LTrim(RTrim(LastName)) FirstLastName, ")
                sb.Append("Role = case ")
                sb.Append("when Role = 'IT' then 'IT' ")
                sb.Append("else dbo.ufn_ToProper(Role) ")
                sb.Append("end, ")
                sb.Append("LocationID, ")
                sb.Append("PhoneExtension ")
            End If

            sb.Append("FROM Users ")
            Sql = sb.ToString

            NonFilterWhereClause = "CompanyID = 'BVI' AND Role <> 'CLIENT' AND StatusCode = 'ACTIVE' AND UserID NOT IN ('eagle', 'eenroller', 'audit', 'bvi', 'falcon')"

            DG.GenerateSQL(Sql, ShowFilter, NonFilterWhereClause, OrderByType, Request, cIndexSess.FilterOnOffState, Request.Form("hdFilterShowHideToggle"))


            ' ___ Eliminate order by clause from recordcount query
            If InfoType = "RecordCount" Then
                Pos = InStr(Sql, "ORDER BY", CompareMethod.Binary)
                If Pos > 0 Then
                    Sql = Sql.Substring(0, Pos - 1)
                End If
            End If

            If InfoType <> "Sql" Then
                QueryPack = cCommon.GetDTExtendedWithQueryPack(Sql, cEnviro.DBHost, "UserManagement")
                If QueryPack.Success Then
                    dt = QueryPack.dt
                Else
                    Throw New Exception("Error #113a: Index GetQueryInfo Info Type: " & InfoType & ". " & QueryPack.TechErrMsg)
                End If
            End If

            If InfoType = "RecordCount" Then
                Coll.Add(dt.Rows(0)(0).Value, "RecordCount")
            ElseIf InfoType = "Data" Then
                Coll.Add(dt, "Data")
            ElseIf InfoType = "Sql" Then
                Coll.Add(Sql, "Sql")
            End If
            Return Coll

        Catch ex As Exception
            Throw New Exception("Error #113: Index GetQueryInfo Info Type: " & InfoType & ". " & ex.Message, ex)
        End Try
    End Function
End Class
