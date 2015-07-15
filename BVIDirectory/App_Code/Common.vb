Imports System.Data.SqlClient
Imports System.IO
Imports System.Data

#Region " Enums "
Public Enum PageMode
    Initial = 1
    Postback = 2
    ReturnFromChild = 3
    CalledByOther = 4
End Enum
Public Enum RequestAction
    CreateNew = 1
    LoadExisting = 2
    SaveNew = 3
    SaveExisting = 4
    SaveNewOrExisting = 5
    ' ConfirmResults = 5
    NoSaveNew = 6
    NoSaveExisting = 7
    ReturnToParent = 8
    [Date] = 9
    Other = 10
End Enum
Public Enum ResponseAction
    DisplayBlank = 1
    DisplayUserInputNew = 2
    DisplayUserInputExisting = 3
    DisplayUserInputNewOrExisting = 4
    DisplayExisting = 5
    ReturnToCallingPage = 6
End Enum

Public Enum StringTreatEnum
    AsIs = 1
    SideQts = 2
    SecApost = 3
    SideQts_SecApost = 4
End Enum
Public Enum ReportTypeEnum
    [Error] = 1
    ErrorNoShutdown = 2
    Timeout = 3
    ProductRegistration = 4
    IAMSRecord = 5
End Enum
Public Enum RoleCatgyEnum
    Other = 1
    Supervisor = 2
    Enroller = 3
End Enum
Public Enum RecordStatusLevelEnum
    EnrollerLevel = 1
    SupervisorLevel = 2
End Enum
Public Enum CallTypeEnum
    New_AcesClientWithEmpID = 4
    New_AcesClientWithoutEmpID = 5
    New_NonAcesClientWithEmpID = 6
    New_NonAcesClientWithoutEmpID = 7
    Existing_AcesClient = 8
    Existing_NonAcesClient = 9
End Enum
Public Enum ProdCreateEditEnum
    Create = 1
    Edit = 2
End Enum
Public Enum TestProdEnum
    Test = 1
    Production = 2
End Enum
Public Enum ProductSourceEnum
    SourceAces = 1
    SourceNative = 2
End Enum

Public Enum ProdMaintTypeEnum
    Standard = 1
    Extended = 2
    AcesSpecial = 3
End Enum
'Public Enum ChangeTypeEnum
'    None = 1
'    Insert = 2
'    Update = 3
'    Delete = 4
'End Enum
Public Enum RelationCodeCatgyEnum
    SingleCode = 1
    MultCodes = 2
    FamilyProdMember = 3
End Enum
#End Region

Public Class Common
    Private cEnviro As Enviro
    Private cDBase As DBase

    Public Sub New()
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        cEnviro = SessionObj("Enviro")
        cDBase = New DBase
    End Sub

    Public Sub New(ByVal Enviro As Enviro)
        cEnviro = Enviro
        cDBase = New DBase
    End Sub

    Public Property EnviroExists() As Boolean
        Get
            If Not IsNothing(cEnviro) Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)

        End Set
    End Property

    Public Sub GenerateError()
        Dim a, b, c As Integer
        a = b / c
    End Sub

    Public Function Right(ByVal Str As String, ByVal Len As Integer) As String
        Return Str.Substring(Str.Length - Len)
    End Function

    Public Sub DropdownFindByValueSelect(ByRef dd As DropDownList, ByVal Value As Object)
        Dim i As Integer
        Dim TestValue As String
        If IsDBNull(Value) OrElse Value = Nothing Then
            TestValue = String.Empty
        Else
            TestValue = Value
        End If
        TestValue = TestValue.ToLower

        For i = 0 To dd.Items.Count - 1
            If dd.Items(i).Value.ToLower = TestValue Then
                dd.Items(i).Selected = True
                Exit For
            End If
        Next
    End Sub

    Public Function Left(ByVal Value As Object, ByVal Length As Integer) As String
        Dim Results As String = String.Empty

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return Results
        Else
            If Value.length <= Length Then
                Return Value.ToString
            Else
                Return Value.Substring(0, Length)
            End If
        End If
    End Function

    Public Function ToProper(ByVal Value As Object) As String
        Dim i, j As Integer
        Dim Box As Object
        Dim Results As String = String.Empty
        Dim CurValue As String = String.Empty
        Dim NewValue As String

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return Results
        Else
            Box = Split(Value, " ")
            For i = 0 To Box.GetUpperBound(0)
                NewValue = String.Empty
                CurValue = Box(i)
                For j = 0 To CurValue.Length - 1
                    If j = 0 Then
                        NewValue = CurValue.Substring(0, 1).ToUpper
                    Else
                        NewValue &= CurValue.Substring(j, 1).ToLower
                    End If
                Next
                Results &= " " & NewValue
            Next
            Results = Trim(Results)
            Return Results
        End If
    End Function

    Public Function ToProperSpecial(ByVal Value As Object) As String
        Dim i, j As Integer
        Dim Box As Object
        Dim Results As String = String.Empty
        Dim NewValue As String

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return Results
        Else
            Box = Split(Value, " ")
            For i = 0 To Box.GetUpperBound(0)
                NewValue = Box(i).Substring(0, 1).ToUpper
                If NewValue.Length > 1 Then
                    NewValue &= Box(i).Substring(1)
                End If
                Results &= " " & NewValue
            Next
            Results = Trim(Results)
            Return Results
        End If
    End Function

    Public Function ToJSAlert(ByVal Value As String) As String
        If Value = Nothing Then
            Value = String.Empty
        End If
        If Value.Length > 0 Then
            Value = Replace(Value, "'", "")
            Value = Replace(Value, "(", "")
            Value = Replace(Value, ")", "")
            Value = Replace(Value, """", "")
            Value = Replace(Value, vbCrLf, " ")
            Value = Replace(Value, Chr(10), " ")
            Value = Replace(Value, Chr(13), " ")
        End If
        Return Value
    End Function

    Public Sub PrintCSVVersionLocal(ByRef dt As DataTable, ByVal FileFullPath As String, ByRef TotalColl As Collection)
        Dim i, j As Integer
        Dim ReportHeader As String
        Dim Sql As New System.Text.StringBuilder
        Dim ColNum As Integer
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim FirstRow As Boolean = True
        Dim RecCount As Integer = 0
        Dim Found As Boolean


        For i = 0 To dt.Rows.Count - 1

            ' ___ Header row.
            If FirstRow Then
                For ColNum = 0 To dt.Columns.Count - 1
                    If ColNum > 0 Then
                        Sql.Append(",")
                    End If
                    Sql.Append("""")
                    Sql.Append(dt.Columns(ColNum).ColumnName)
                    Sql.Append("""")
                Next
                Sql.Append(vbCrLf)
                FirstRow = False
            End If

            ' ___ Data rows.
            For ColNum = 0 To dt.Columns.Count - 1
                If ColNum > 0 Then
                    Sql.Append(",")
                End If
                Sql.Append("""")
                Try
                    Sql.Append(dt.Rows(i)(ColNum).ToJSParm)
                Catch
                    Sql.Append(dt.Rows(i)(ColNum))
                End Try

                Sql.Append("""")
            Next
            Sql.Append(vbCrLf)

        Next

        ' ___ Totals
        If Not TotalColl Is Nothing Then
            For ColNum = 0 To dt.Columns.Count - 1
                If ColNum = 0 Then
                    Sql.Append("""")
                    Sql.Append("TOTAL")
                    Sql.Append("""")
                Else
                    Sql.Append(",")
                    For j = 1 To TotalColl.Count
                        Found = False
                        If TotalColl(j).ItemName = dt.Columns(ColNum).ColumnName Then
                            Sql.Append("""")
                            Sql.Append(TotalColl(j).Value)
                            Sql.Append("""")
                            Found = True
                            Exit For
                        End If
                    Next
                End If
            Next
            Sql.Append(vbCrLf)
        End If


        If FileFullPath = Nothing Then
            FileFullPath = cEnviro.ApplicationPath & "\csv.csv"
        End If

        If File.Exists(FileFullPath) Then
            File.Delete(FileFullPath)
        End If
        fs = File.Create(FileFullPath)
        fs.Close()

        sw = New StreamWriter(FileFullPath)
        sw.Write(Sql.ToString)
        sw.Close()
        sw.Close()
        fs.Close()

    End Sub

    Public Function GetDownloadPath(ByVal Page As Page) As Collection
        Dim PathsColl As New Collection
        Dim DirPath As String
        Dim FileName As String
        Dim FullPath As String

        DirPath = Page.Request.ServerVariables("APPL_PHYSICAL_PATH") & "TempData\"
        'FileName = "Rpt" & DateDiff("s", "1/1/2000", Now) & "a" & CInt(Rnd() * 1000) & ".csv"
        FileName = "Rpt_" & cEnviro.LoggedInUserID & "_" & Date.Now.ToUniversalTime.AddHours(-5).ToString("yyyyMMdd_HHmmss_fff") & ".csv"
        FullPath = DirPath & FileName
        PathsColl.Add(DirPath, "DirPath")
        PathsColl.Add(FullPath, "AbsPath")
        PathsColl.Add("./TempData/" & FileName, "RelPath")
        Return PathsColl
    End Function

    Public Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String)
        Dim Report As New Report
        Report.SendEmail(SendTo, From, cc, Subject, TextBody)
    End Sub

    Public Function GetPageCaption() As String
        Return "BVI Directory v" & cEnviro.VersionNumber
    End Function

#Region " Data "
    Public Function ConvertToExtendedTable(ByRef dt As DataTable) As DataTable
        Return cDBase.GetDTExtended(dt)
    End Function
#End Region

#Region " New GetData "
    Public Sub ExecuteNonQuery(ByVal Sql As String)
        ExecuteNonQueryMaster(Sql, cEnviro.DBHost, cEnviro.DefaultDatabase, False)
    End Sub

    Public Sub ExecuteNonQuery(ByVal DBHost As String, ByVal Database As String, ByVal Sql As String)
        ExecuteNonQueryMaster(Sql, DBHost, Database, False)
    End Sub

    Public Function ExecuteNonQueryWithQuerypack(ByVal Sql As String) As DBase.QueryPack
        Dim Querypack As DBase.QueryPack
        Querypack = ExecuteNonQueryMaster(Sql, cEnviro.DBHost, cEnviro.DefaultDatabase, True)
        Return Querypack
    End Function

    Public Function ExecuteNonQueryWithQuerypack(ByVal DBHost As String, ByVal Database As String, ByVal Sql As String) As DBase.QueryPack
        Dim Querypack As New DBase.QueryPack
        Querypack = ExecuteNonQueryMaster(Sql, DBHost, Database, True)
        Return Querypack
    End Function

    Public Function ExecuteNonQueryMaster(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal WithQuerypack As Boolean) As DBase.QueryPack
        Dim Querypack As New DBase.QueryPack

        Try
            Dim SqlConnection1 As New SqlClient.SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
            SqlConnection1.Open()
            Dim SqlCmd As New System.Data.SqlClient.SqlCommand(Sql, SqlConnection1)
            SqlCmd.CommandType = System.Data.CommandType.Text
            SqlCmd.ExecuteNonQuery()
            SqlCmd.Dispose()
            SqlConnection1.Close()
            Querypack.Success = True
        Catch ex As Exception

            If Not WithQuerypack Then
                Throw New Exception("Error #CM2203: Common ExecuteNonQueryMaster.~Sql: " & Sql & "~DBHost: " & DBHost & "~Database: " & Database & "~Error message: " & ex.Message)
            End If

            Querypack.Success = False
            Querypack.TechErrMsg = ex.Message
        End Try
        Return Querypack
    End Function


    ' ___ GetDT
    Public Function GetDT(ByVal Sql As String) As DataTable
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, cEnviro.DBHost, cEnviro.DefaultDatabase, False, False)
        Return QueryPack.dt
    End Function

    Public Function GetDT(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String) As DataTable
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, False, False)
        Return QueryPack.dt
    End Function

    Public Function GetDTWithQueryPack(ByVal Sql As String) As DBase.QueryPack
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, cEnviro.DBHost, cEnviro.DefaultDatabase, False, True)
        Return QueryPack
    End Function

    Public Function GetDTWithQueryPack(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String) As DBase.QueryPack
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, False, True)
        Return QueryPack
    End Function

    ' ___ GetDTExtended
    Public Function GetDTExtended(ByVal Sql As String) As DataTable
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, cEnviro.DBHost, cEnviro.DefaultDatabase, True, False)
        Return QueryPack.dt
    End Function

    Public Function GetDTExtended(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String) As DataTable
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, True, False)
        Return QueryPack.dt
    End Function

    Public Function GetDTExtendedWithQueryPack(ByVal Sql As String) As DBase.QueryPack
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, cEnviro.DBHost, cEnviro.DefaultDatabase, True, True)
        Return QueryPack
    End Function

    Public Function GetDTExtendedWithQueryPack(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String) As DBase.QueryPack
        Dim QueryPack As New DBase.QueryPack
        QueryPack = GetDTMaster(Sql, DBHost, Database, True, True)
        Return QueryPack
    End Function

    Public Function GetDTMaster(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal ExtendedTbl As Boolean, ByVal WithQuerypack As Boolean) As DBase.QueryPack
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable
        Dim QueryPack As New DBase.QueryPack

        Dim SqlCmd As New SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.CommandTimeout = 90
        SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
        DataAdapter = New SqlDataAdapter(SqlCmd)
        Try
            DataAdapter.Fill(dt)
            If ExtendedTbl Then
                dt = cDBase.GetDTExtended(dt)
            End If
            QueryPack.Success = True
            QueryPack.dt = dt
        Catch ex As Exception

            If Not WithQuerypack Then
                Throw New Exception("Error #CM2202: Common GetDTMaster.~Sql: " & Sql & "~DBHost: " & DBHost & "~Database: " & Database & "~Error message: " & ex.Message)
            End If

            QueryPack.Success = False
            QueryPack.TechErrMsg = ex.Message
        End Try

        DataAdapter.Dispose()
        SqlCmd.Dispose()
        SqlCmd.Connection.Close()
        Return QueryPack
    End Function

    'Public Function GetDTExtendedWithQueryPack(ByVal Sql As String, ByVal DBHost As String, ByVal Database As String, ByVal ExtendedTbl As Boolean) As DBase.QueryPack
    '    Dim DataAdapter As SqlDataAdapter
    '    Dim dt As New DataTable
    '    Dim QueryPack As New DBase.QueryPack

    '    Dim SqlCmd As New SqlCommand(Sql)
    '    SqlCmd.CommandType = CommandType.Text
    '    SqlCmd.Connection = New SqlConnection(cEnviro.GetConnectionString(DBHost, Database))
    '    DataAdapter = New SqlDataAdapter(SqlCmd)
    '    Try
    '        DataAdapter.Fill(dt)
    '        If ExtendedTbl Then
    '            dt = cDBase.GetDTExtended(dt)
    '        End If
    '        QueryPack.Success = True
    '        QueryPack.dt = dt
    '    Catch ex As Exception
    '        QueryPack.Success = False
    '        QueryPack.TechErrMsg = ex.Message
    '    End Try

    '    DataAdapter.Dispose()
    '    SqlCmd.Dispose()
    '    Return QueryPack
    'End Function
#End Region

#Region " Page handling "
    Public Function GetPageMode(ByVal Page As Page, ByVal Sess As PageSession) As PageMode
        Dim PageMode As PageMode

        If Page.IsPostBack AndAlso Page.Request.Form("__EVENTTARGET") = "" Then
            PageMode = PageMode.Postback
        Else
            Select Case Page.Request.QueryString("CalledBy")
                Case "Child"
                    PageMode = PageMode.ReturnFromChild
                Case "Other"
                    If Sess.PageInitiallyLoaded Then
                        PageMode = PageMode.CalledByOther
                    Else
                        PageMode = PageMode.Initial
                        Sess.PageInitiallyLoaded = True
                    End If
                Case Else
                    PageMode = PageMode.Initial
                    Sess.PageInitiallyLoaded = True
            End Select
        End If
        Return PageMode
    End Function

    Public Function GetRequestAction(ByVal Page As Page) As RequestAction
        Dim ActionType As String
        Dim RequestAction As RequestAction
        Dim hdResponseAction As String
        Dim hdAction As String
        Dim CallType As String

        If Page.Request.QueryString("CallType") = Nothing OrElse Page.Request.QueryString("CallType") = String.Empty Then
            CallType = String.Empty
        Else
            CallType = Page.Request.QueryString("CallType")
            CallType = CallType.ToLower
        End If

        If Page.Request.Form("hdAction") = Nothing OrElse Page.Request.Form("hdAction") = String.Empty Then
            hdAction = String.Empty
        Else
            hdAction = Page.Request.Form("hdAction")
            hdAction = hdAction.ToLower
        End If

        If Not Page.IsPostBack Then
            ActionType = "record"
        Else
            If hdAction = "update" Then
                ActionType = "record"
            ElseIf hdAction = "return" Then
                ActionType = "return"
            ElseIf hdAction = "confirmation" Then
                ActionType = "confirmation"
            ElseIf hdAction = "clientselectionchanged" Then
                ActionType = "clientselectionchanged"
            End If
        End If

        Select Case ActionType
            Case "return"
                RequestAction = RequestAction.ReturnToParent

            Case "record"
                If Not Page.IsPostBack Then
                    Select Case CallType
                        Case "new", ""
                            RequestAction = RequestAction.CreateNew
                        Case "existing"
                            RequestAction = RequestAction.LoadExisting
                    End Select
                Else
                    Select Case Page.Request.Form("hdResponseAction")
                        Case ResponseAction.DisplayBlank.ToString
                            RequestAction = RequestAction.SaveNew
                        Case ResponseAction.DisplayExisting.ToString
                            RequestAction = RequestAction.SaveExisting
                        Case ResponseAction.DisplayUserInputNew.ToString
                            RequestAction = RequestAction.SaveNew
                        Case ResponseAction.DisplayUserInputExisting.ToString
                            RequestAction = RequestAction.SaveExisting
                        Case ResponseAction.DisplayUserInputNewOrExisting.ToString
                            RequestAction = RequestAction.SaveNewOrExisting
                    End Select
                End If

            Case "confirmation"
                hdResponseAction = Page.Request.Form("hdResponseAction")
                If hdResponseAction = "DisplayBlank" Or hdResponseAction = "DisplayUserInputNew" Then
                    If Page.Request.Form("hdConfirm") = "yes" Then
                        RequestAction = RequestAction.SaveNew
                    Else
                        RequestAction = RequestAction.NoSaveNew
                    End If
                ElseIf hdResponseAction = "DisplayExisting" Or hdResponseAction = "DisplayUserInputExisting" Then
                    If Page.Request.Form("hdConfirm") = "yes" Then
                        RequestAction = RequestAction.SaveExisting
                    Else
                        RequestAction = RequestAction.NoSaveExisting
                    End If
                End If

            Case "clientselectionchanged"
                RequestAction = RequestAction.Other

        End Select

        Return RequestAction
    End Function

    Public Function GetResponseActionFromRequestActionOther(ByRef Page As Page) As ResponseAction
        Select Case Page.Request("hdResponseAction").ToString
            Case ResponseAction.DisplayBlank.ToString
                Return ResponseAction.DisplayUserInputNew
            Case ResponseAction.DisplayUserInputNew.ToString
                Return ResponseAction.DisplayUserInputNew
            Case ResponseAction.DisplayUserInputExisting.ToString
                Return ResponseAction.DisplayUserInputExisting
            Case ResponseAction.DisplayUserInputNewOrExisting.ToString
                Return ResponseAction.DisplayUserInputNewOrExisting
            Case ResponseAction.DisplayExisting.ToString
                Return ResponseAction.DisplayUserInputExisting
            Case ResponseAction.ReturnToCallingPage.ToString
                Return ResponseAction.ReturnToCallingPage
        End Select
    End Function
#End Region

#Region " In handlers "
    Public Function StrInHandler(ByVal Input As Object) As Object
        Dim Output As Object

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf (Not IsNumeric(Input)) AndAlso Input = Nothing Then
            Return String.Empty
            'ElseIf (Not IsDate(Input)) AndAlso Input.length = 0 Then
            '    Return String.Empty
        Else
            Output = Input
            If Input = Nothing Then
                Return String.Empty
            End If
            Return Output
        End If
    End Function

    Public Function DateInHandler(ByVal Input As Object) As Object
        ' 12/31/2399
        Dim Output As Object
        Output = Input

        If IsDBNull(Input) Then
            Return String.Empty
        ElseIf Input = "01/01/1900" Then
            Return String.Empty
        ElseIf Input = "01/01/1950" Then
            Return String.Empty
        Else
            Return Output
        End If
    End Function

    Public Function NumInHandler(ByVal Input As Object, ByVal NullAsZero As Boolean) As Object
        If IsDBNull(Input) Then
            If NullAsZero Then
                Return 0
            Else
                Return String.Empty
            End If
        Else
            Return Input
        End If
    End Function

    Public Function GuidInHandler(ByVal Input As Object) As Object
        If IsDBNull(Input) Then
            Return String.Empty
        Else
            Return Input.ToString
        End If
    End Function


    Public Function StrXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        Dim Output As Object
        Dim ReturnNull As Boolean

        If IsDBNull(Input) Then
            ReturnNull = True
        ElseIf (Not IsNumeric(Input)) AndAlso Input = Nothing Then
            ReturnNull = True
        Else
            Output = Replace(Input, "~", "'")
            If Output = Nothing Then
                ReturnNull = True
            End If
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return String.Empty
            End If
        Else
            Return Output
        End If

    End Function
    Public Function DateXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        ' 12/31/2399
        Dim Output As Object
        Dim ReturnNull As Boolean
        Output = Input

        If IsDBNull(Input) OrElse Input = Nothing Then
            ReturnNull = True
        Else
            Output = Input
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return "1/1/1950"
            End If
        Else
            Return Output
        End If
    End Function

    Public Function NumXferHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
        Dim Output As Object
        Dim ReturnNull As Boolean

        If IsDBNull(Input) Then
            ReturnNull = True
        Else
            Output = Input
        End If

        If ReturnNull Then
            If AllowNull Then
                Return DBNull.Value
            Else
                Return 0
            End If
        Else
            Return Output
        End If

    End Function
#End Region

#Region " Out handlers"
    'Public Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
    '    Dim ReturnNull As Boolean
    '    Dim Output As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input.length > 0 Then
    '        Output = Replace(Input, "'", "~")
    '    Else
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    End If

    '    If ReturnNull Then
    '        Return "null"
    '    Else
    '        If AddSingleQuotes Then
    '            Return "'" & Output & "'"
    '        Else
    '            Return Output
    '        End If
    '    End If
    'End Function

    'Public Function OrigStrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
    '    Dim ReturnNull As Boolean
    '    Dim Output As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input.length > 0 Then
    '        'Output = Replace(Input, "'", "~")

    '        If AddSingleQuotes Then
    '            Output = Replace(Input, "'", "''")
    '        Else
    '            Output = Input
    '        End If

    '    Else
    '        If AllowNull Then
    '            ReturnNull = True
    '        Else
    '            Output = String.Empty
    '        End If
    '    End If

    '    If ReturnNull Then
    '        Return "null"
    '    Else
    '        If AddSingleQuotes Then
    '            Return "'" & Output & "'"
    '        Else
    '            Return Output
    '        End If
    '    End If
    'End Function

    'Public Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal StringTreat As StringTreatEnum) As Object
    '    Dim Output As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = String.Empty
    '        End If
    '    ElseIf Input.length > 0 Then
    '        Output = Input
    '        If (StringTreat = StringTreatEnum.SecApost) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
    '            Output = Replace(Output, "'", "''")
    '        End If
    '        If (StringTreat = StringTreatEnum.SideQts) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
    '            Output = "'" & Output & "'"
    '        End If
    '    Else
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = String.Empty
    '            If (StringTreat = StringTreatEnum.SideQts) Or (StringTreat = StringTreatEnum.SideQts_SecApost) Then
    '                Output = "'" & Output & "'"
    '            End If
    '        End If
    '    End If
    '    Return Output
    'End Function


    Public Function StrOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean, ByVal StringTreat As StringTreatEnum) As Object
        Dim Output As String

        Try

            ' ___ Output, adjusting for AllowNull
            If IsDBNull(Input) Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = String.Empty
                End If
            ElseIf Input Is Nothing Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = String.Empty
                End If
            Else
                Try
                    Output = Input
                Catch
                    If AllowNull Then
                        Output = "null"
                    Else
                        Output = String.Empty
                    End If
                End Try
            End If

            ' ___ Apply string treatment
            If Output <> "null" Then
                Select Case StringTreat
                    Case StringTreatEnum.AsIs
                        ' no action
                    Case StringTreatEnum.SecApost
                        Output = Replace(Output, "'", "''")
                    Case StringTreatEnum.SideQts
                        Output = "'" & Output & "'"
                    Case StringTreatEnum.SideQts_SecApost
                        Output = Replace(Output, "'", "''")
                        Output = "'" & Output & "'"
                End Select
            End If

            Return Output

        Catch ex As Exception
            Throw New Exception("Error #CM2220: Common StrOutHandler. " & ex.Message, ex)
        End Try
    End Function

    'Public Function NumOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean) As String
    '    Dim i As Integer
    '    Dim Output As String = String.Empty
    '    Dim Allowable As String
    '    Dim Working As String

    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = "0"
    '        End If
    '    ElseIf Input Is Nothing Then
    '        If AllowNull Then
    '            Output = "null"
    '        Else
    '            Output = "0"
    '        End If
    '    Else
    '        Try
    '            Working = Input
    '            Allowable = "-.0123456789"
    '            For i = 0 To Working.Length - 1
    '                If InStr(Allowable, Working.Substring(i, 1)) > 0 Then
    '                    Output &= Working.Substring(i, 1)
    '                End If
    '            Next
    '        Catch ex As Exception
    '            If AllowNull Then
    '                Output = "null"
    '            Else
    '                Output = "0"
    '            End If
    '        End Try
    '    End If

    '    Return Output
    'End Function

    Public Function NumOutHandler(ByRef Input As Object, ByVal AllowNull As Boolean) As String
        Dim i As Integer
        Dim Output As String = String.Empty
        Dim Allowable As String
        Dim Working As String

        Try

            If IsDBNull(Input) Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = "0"
                End If
            ElseIf Input Is Nothing Then
                If AllowNull Then
                    Output = "null"
                Else
                    Output = "0"
                End If
            Else

                Try
                    Working = Input
                    Allowable = "-.0123456789"
                    For i = 0 To Working.Length - 1
                        If InStr(Allowable, Working.Substring(i, 1)) > 0 Then
                            Output &= Working.Substring(i, 1)
                        End If
                    Next
                    If Output.Length = 0 Then
                        If AllowNull Then
                            Output = "null"
                        Else
                            Output = "0"
                        End If
                    End If
                Catch
                    If AllowNull Then
                        Output = "null"
                    Else
                        Output = "0"
                    End If
                End Try
            End If

        Catch
            If AllowNull Then
                Output = "null"
            Else
                Output = "0"
            End If
        End Try

        Return Output
    End Function

    Public Function DateOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim ReturnNull As Boolean
        Dim Output As Object

        If IsDBNull(Input) OrElse Input = Nothing Then
            If AllowNull Then
                ReturnNull = True
            Else
                Output = "01/01/1950"
            End If
        Else
            Output = Input
        End If

        If ReturnNull Then
            Return "null"
        Else
            If AddSingleQuotes Then
                Return "'" & Output & "'"
            Else
                Return Output
            End If
        End If
    End Function

    Public Function DateOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, ByVal FormatString As String, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim ReturnNull As Boolean
        Dim Output As Object

        If IsDBNull(Input) OrElse Input = Nothing Then
            If AllowNull Then
                ReturnNull = True
            Else
                Output = "01/01/1950"
            End If
        Else
            Output = Input
        End If

        If (Not ReturnNull) And IsDate(Output) Then
            Output = CType(Output, System.DateTime).ToString(FormatString)
        End If

        If ReturnNull Then
            Return "null"
        Else
            If AddSingleQuotes Then
                Return "'" & Output & "'"
            Else
                Return Output
            End If
        End If
    End Function

    Public Function PhoneOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, Optional ByVal AddSingleQuotes As Boolean = False) As Object
        Dim i As Integer
        Dim Output As String = String.Empty
        Dim Working As String
        Working = StrOutHandler(Input, AllowNull, StringTreatEnum.SideQts)

        If Working = "null" Or Working = String.Empty Then
        Else
            If Working.Length >= 10 Then
                For i = 0 To Working.Length - 1
                    If IsNumeric(Working.Substring(i, 1)) Then
                        Output &= Working.Substring(i, 1)
                    End If
                Next
            End If
        End If

        If Output.Length = 10 Then
            Output = InsertAt(Output, "(", 1)
            Output = InsertAt(Output, ") ", 5)
            Output = InsertAt(Output, "-", 10)
        Else
            Output = Input
        End If

        If AddSingleQuotes Then
            Output = "'" & Output & "'"
        End If

        Return Output
    End Function

    'Public Function BitOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean) As Object
    '    If IsDBNull(Input) Then
    '        If AllowNull Then
    '            Return "null"
    '        Else
    '            Return 0
    '        End If
    '    Else
    '        If CType(Input, Boolean) Then
    '            Return 1
    '        Else
    '            Return 0
    '        End If
    '    End If
    'End Function

    Public Function BitOutHandler(ByVal Input As Object, ByVal AllowNull As Boolean, ByVal AddSingleQuoteToNull As Boolean) As Object
        Try
            If IsDBNull(Input) Then
                If AllowNull Then
                    Return "null"
                Else
                    Return 0
                End If
            ElseIf Input Is Nothing Then
                If AllowNull Then
                    Return "null"
                Else
                    Return 0
                End If
            ElseIf Input = String.Empty Then
                If AllowNull Then
                    Return "null"
                Else
                    Return 0
                End If

            Else
                If CType(Input, Boolean) Then
                    Return 1
                Else
                    Return 0
                End If
            End If
        Catch
            If AllowNull Then
                Return "null"
            Else
                Return 0
            End If
        End Try
    End Function


#End Region

#Region " Validate "

    Public Function IsStrEqual(ByVal FirstValue As String, ByVal SecondValue As String, Optional ByVal IgnoreCase As Boolean = True) As Boolean
        Dim Output As Integer
        Dim Results As Boolean
        Output = String.Compare(Trim(FirstValue), Trim(SecondValue), IgnoreCase)
        If Output = 0 Then
            Results = True
        End If
        Return Results
    End Function

    Public Function InList(ByVal SearchFor As String, ByVal SearchInList As String, Optional ByVal IgnoreCase As Boolean = True) As Boolean
        Dim i As Integer
        Dim Box() As String

        If SearchInList Is Nothing OrElse SearchInList.Length = 0 Then
            Return False
        Else
            Box = Split(SearchInList, ",")
            For i = 0 To Box.GetUpperBound(0)
                If IsStrEqual(SearchFor, Trim(Box(i)), IgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End If
    End Function

    Public Function IsBlank(ByVal Value As Object) As Boolean
        If IsDBNull(Value) Then
            Return True
        ElseIf Value = Nothing Then
            Return True
        Else
            If Value.length = 0 Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Sub ValidateStringField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal MinLength As Integer, ByVal ErrMsg As String)
        If Value.length < MinLength Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateStringField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal MinLength As Integer, ByVal MaxLength As Integer, ByVal ErrMsg As String)
        If Value.length < MinLength Or Value.length > MaxLength Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateApostrophe(ByRef ErrColl As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If InStr(Value, "'") > 0 Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateNumericField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim PassTest As Boolean
        If IsDBNull(Value) OrElse Value.Length = 0 Then
            If AllowNull Then
                PassTest = True
            Else
                PassTest = False
            End If
        Else
            If IsNumeric(Value) Then
                PassTest = True
            Else
                PassTest = False
            End If
        End If

        If Not PassTest Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateDateTime(ByRef ErrColl As Collection, ByVal Input As Object, ByVal ValidateDatePart As Boolean, ByVal ValidateTimePart As Boolean, ByVal ErrMsg As String)
        Dim DatePart As String
        Dim TimePart As String
        Dim DateIsValid As Boolean = True
        Dim TimeIsValid As Boolean = True
        Dim ReportError As Boolean

        If ValidateDatePart Then
            Try
                DatePart = CType(Input, System.DateTime).ToString("MM/dd/yyyy")
                If Not IsDate(DatePart) Then
                    DateIsValid = False
                End If
            Catch
                DateIsValid = False
            End Try
        End If


        If ValidateTimePart Then
            Try
                TimePart = CType(Input, System.DateTime).ToString("hh:mm tt")
                If Not IsDate(CType(Date.Now.ToString("MM/dd/yyyy") & " " & TimePart, System.DateTime)) Then
                    TimeIsValid = False
                End If
            Catch
                TimeIsValid = False
            End Try
        End If

        If ValidateDatePart And Not ValidateTimePart Then
            If Not DateIsValid Then
                ReportError = True
            End If
        ElseIf Not ValidateDatePart And ValidateTimePart Then
            If Not TimeIsValid Then
                ReportError = True
            End If
            If ValidateDatePart And ValidateTimePart Then
                If DateIsValid And Not TimeIsValid Then
                    ReportError = True
                ElseIf Not DateIsValid And TimeIsValid Then
                    ReportError = True
                ElseIf Not DateIsValid And Not TimeIsValid Then
                    ReportError = True
                End If
            End If
        End If

        If ReportError Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateRadio(ByRef ErrColl As Collection, ByVal SelectedIndex As Integer, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        If (SelectedIndex < 0) AndAlso (Not AllowNull) Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateNumericRange(ByRef ErrColl As Collection, ByVal Value As Object, ByVal Min As Integer, ByVal Max As Integer, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim PassTest As Boolean
        If IsDBNull(Value) OrElse Value.Length = 0 Then
            If AllowNull Then
                PassTest = True
            Else
                PassTest = False
            End If
        Else
            If IsNumeric(Value) Then
                If Value >= Min AndAlso Value <= Max Then
                    PassTest = True
                Else
                    PassTest = False
                End If
            End If
        End If

        If Not PassTest Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateDateField(ByRef ErrColl As Collection, ByVal Value As Object, ByVal AllowNull As Boolean, ByVal ErrMsg As String)
        Dim Valid As Boolean
        If IsDBNull(Value) OrElse Value = Nothing Then
            If AllowNull Then
                Valid = True
            Else
                Valid = False
            End If
        ElseIf IsDate(Value) Then
            Valid = True
        Else
            Valid = False
        End If
        If Not Valid Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Function IsValidPhoneNumber(ByVal Value As Object) As Boolean
        Dim i As Integer
        Dim NumCount As Integer

        If IsDBNull(Value) OrElse Value = Nothing Then
            Return False
        End If

        If Value.length >= 10 Then
            For i = 0 To Value.Length - 1
                If IsNumeric(Value.Substring(i, 1)) Then
                    NumCount += 1
                End If
            Next
        End If

        If NumCount = 10 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub ValidatePhoneNumber(ByRef Errcoll As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If Not IsValidPhoneNumber(Value) Then
            If Errcoll.Count = 0 Then
                Errcoll.Add(ErrMsg)
            Else
                Errcoll.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateEmailAddress(ByRef ErrColl As Collection, ByVal Value As Object, ByVal ErrMsg As String)
        If Not IsValidEmailAddress(Value) Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Function IsValidEmailAddress(ByVal Value As Object) As Boolean
        Dim OKSoFar As Boolean = True
        Const InvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
        Dim i As Integer
        Dim Num As Integer
        Dim DotPos As Integer
        Dim Part2 As String

        ' ___ Check for null or empty value
        If IsDBNull(Value) OrElse Value = Nothing Then
            OKSoFar = False
        End If

        ' ___ Check for minimum length
        If OKSoFar Then
            If Value.Length < 5 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for a double quote
        If OKSoFar Then
            OKSoFar = Not InStr(1, Value, Chr(34)) > 0  'Check to see if there is a double quote
        End If

        ' ___ Check for consecutive dots
        If OKSoFar Then
            OKSoFar = Not InStr(1, Value, "..") > 0
        End If

        ' ___ Check for invalid characters
        If OKSoFar Then
            For i = 0 To InvalidChars.Length - 1
                If InStr(1, Value, InvalidChars.Substring(i, 1)) > 0 Then
                    OKSoFar = False
                    Exit For
                End If
            Next
        End If

        ' ___ Check for number of @ symbols
        If OKSoFar Then
            For i = 0 To Value.Length - 1
                If InStr(Value.Substring(i, 1), "@") > 0 Then
                    Num += 1
                End If
            Next
            If Num > 1 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for the @ symbol in starting before the third position
        If OKSoFar Then
            If InStr(Value, "@") < 2 Then
                OKSoFar = False
            End If
        End If

        ' ___ Check for number of dots
        If OKSoFar Then
            Num = 0
            Part2 = Value.substring(InStr(Value, "@"))
            For i = 0 To Part2.Length - 1
                If InStr(Part2.Substring(i, 1), ".") > 0 Then
                    Num += 1
                End If
            Next
            If Num > 1 Then
                OKSoFar = False
            End If
        End If

        ' ___ Dot is present and not immediately after ampersand and not at end. 
        '___  Dot separated from ampersand by at least one character
        If OKSoFar Then
            DotPos = InStr(Part2, ".")
            If DotPos < 2 Or DotPos = Part2.Length Then
                OKSoFar = False
            End If
        End If

        Return OKSoFar
    End Function

    Public Sub ValidateDropDown(ByRef ErrColl As Collection, ByRef dd As DropDownList, ByVal MinSelectedIndex As Integer, ByVal ErrMsg As String)
        If dd.SelectedIndex < MinSelectedIndex Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateDropDownSelect0(ByRef ErrColl As Collection, ByRef dd As DropDownList, ByVal ErrMsg As String)
        If dd.SelectedIndex > 0 Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateCheckbox(ByRef ErrColl As Collection, ByRef chkBox As CheckBox, ByVal ValidState As Integer, ByVal ErrMsg As String)
        Dim IsValid As Boolean = True
        If ValidState = 0 AndAlso Not chkBox.Checked Then
            IsValid = False
        ElseIf ValidState = 1 AndAlso Not chkBox.Checked Then
            IsValid = False
        End If
        If Not IsValid Then
            If ErrColl.Count = 0 Then
                ErrColl.Add(ErrMsg)
            Else
                ErrColl.Add(", " & ErrMsg)
            End If
        End If
    End Sub

    Public Sub ValidateErrorOnly(ByRef ErrColl As Collection, ByVal ErrMsg As String)
        If ErrColl.Count = 0 Then
            ErrColl.Add(ErrMsg)
        Else
            ErrColl.Add(", " & ErrMsg)
        End If
    End Sub

    Public Function ErrCollToString(ByRef ErrColl As Collection, ByVal Intro As String) As String
        Dim sb As New System.Text.StringBuilder
        Dim i As Integer
        If ErrColl.Count > 0 Then
            For i = 1 To ErrColl.Count
                sb.Append(ErrColl(i))
            Next
        End If
        Return Intro & " " & sb.ToString & "."
    End Function
#End Region

#Region " This to that "
    Public Function BitToRadio(ByVal Value As Object, ByVal TrueIndex As Integer, ByVal AllowNoneSelected As Boolean) As Integer
        Dim FalseIndex As Integer
        FalseIndex = System.Math.Abs(TrueIndex - 1)

        If IsDBNull(Value) Then
            If AllowNoneSelected Then
                Return -1
            Else
                Return FalseIndex
            End If
        Else
            If Value Then
                Return TrueIndex
            Else
                Return FalseIndex
            End If
        End If
    End Function

    Public Function BitToString(ByVal Value As Object, ByVal TrueString As String, ByVal FalseString As String, ByVal AllowNull As Boolean) As String
        If IsDBNull(Value) Then
            If AllowNull Then
                Return String.Empty
            Else
                Return FalseString
            End If
        End If
        If Value Then
            Return TrueString
        Else
            Return FalseString
        End If
    End Function

    Public Function BitToInt(ByVal Value As Object) As Integer
        If IsDBNull(Value) Then
            Return 0
        Else
            If Value Then
                Return 1
            Else
                Return 0
            End If
        End If
    End Function

    Public Function ChkToInd(ByVal chkBox As CheckBox) As Integer
        If chkBox.Checked Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Public Sub IndToChk(ByVal Ind As Object, ByVal chkBox As CheckBox)
        If IsDBNull(Ind) Then
            chkBox.Checked = False
        Else
            If Ind Then
                chkBox.Checked = True
            Else
                chkBox.Checked = False
            End If
        End If
    End Sub
#End Region

#Region " Everything else "
    Public Function GetServerDateTime() As DateTime
        Return Date.Now.ToUniversalTime.AddHours(-5)
    End Function

    Public Function ConditionStringForHTML(ByVal Value As Object) As String
        Dim Results As String
        If IsDBNull(Value) Then
            Results = String.Empty
        Else
            Results = Value.ToString
        End If
        Results = Replace(Results, Chr(10).ToString, "<br />")
        Return Results
    End Function

    Public Function GetRightsStr(ByRef dt As DataTable) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        If dt.Rows.Count = 0 Then
            Return String.Empty
        Else
            For i = 0 To dt.Rows.Count - 1
                sb.Append("|" & StrInHandler(dt.Rows(i)("RightCd")))
            Next
            Return sb.ToString.Substring(1)
        End If
    End Function

    Public Function InsertAt(ByVal Value As String, ByVal InsChar As String, ByVal Pos As Integer) As String
        Dim ValuePos As Integer = 1
        Dim Output As String = String.Empty
        Dim OutputPos As Integer = 1
        Do
            If OutputPos = Pos Then
                Output &= InsChar
                OutputPos += 1
            Else
                Output &= Value.Substring(ValuePos - 1, 1)
                ValuePos += 1
                OutputPos += 1
                If ValuePos > Value.Length Then
                    Exit Do
                End If
            End If
        Loop
        Return Output
    End Function

    Public Function ToNull(ByVal Input As Object) As Object
        If IsDBNull(Input) Then
            Return DBNull.Value
        ElseIf Input Is Nothing Then
            Return DBNull.Value
        ElseIf Input.length = 0 Then
            Return DBNull.Value
        Else
            Return Input
        End If
    End Function

    Public Function IsBVIDate(ByVal Input As Object) As Boolean
        If IsDBNull(Input) Then
            Return False
        ElseIf Input = Nothing Then
            Return False
        ElseIf Input = "01/01/1950" Then
            Return False
        ElseIf Input.ToString = String.Empty Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function GetCurRightsHidden(ByVal RightsColl As Collection) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        For i = 1 To RightsColl.Count
            sb.Append(RightsColl(i) & "|")
        Next
        sb.Length -= 1
        Return "<input type='hidden' id='currentrights' name='currentrights' value=""" & sb.ToString & """>"
    End Function

    Public Function GetCurRightsHidden(ByVal RightsStr As String) As String
        Return "<input type='hidden'id='currentrights' name='currentrights' value=""" & RightsStr & """ > "
    End Function

    Public Function GetCurRightsAndTopicsHidden(ByVal RightsColl As Collection) As String
        Dim i As Integer

        Dim sb As New System.Text.StringBuilder
        For i = 1 To RightsColl.Count
            sb.Append(RightsColl(i) & "|")
        Next
        sb.Length -= 1
        Return "<input type='hidden' id='currentrights' name='currentrights' value=""" & sb.ToString & """><input type='hidden' id='currenttopics' name = 'currenttopics' value=""PUB|" & cEnviro.LogInRoleCatgy.ToString & """>"
    End Function

    Public Function GetCurRightsAndTopicsHidden(ByVal RightsStr As String) As String
        Return "<input type='hidden'id='currentrights' name='currentrights' value=""" & RightsStr & """ > "
    End Function

    Public Function DTSort(ByRef dt As DataTable, ByVal Filter As String, ByVal Sort As String, ByVal Ascending As Boolean) As DataTable
        'strExpr = "id > 5"
        ' Sort descending by CompanyName column.
        'strSort = "name DESC"
        ' Use the Select method to find all rows matching the filter.

        ' To filter out the null first row in a DS DT:
        ' Filter = "CUST_ID > ''"
        '.FilterOn("Holiday = 'xmas'")

        Dim i As Integer
        Dim Row As Integer
        'Dim Col As Integer
        Dim SortedRows() As DataRow
        Dim NewDT As New DataTable
        Dim dr As DataRow
        Dim CompoundSort As Boolean

        Try

            ' ___ Use the Select method to find all rows matching the filter.
            If Sort = Nothing Then
                SortedRows = dt.Select(Filter)
            Else
                If InStr(Sort, " ") > 0 Then
                    CompoundSort = True
                End If
                If CompoundSort Then
                    SortedRows = dt.Select(Filter, Sort)
                Else
                    If Ascending Then
                        SortedRows = dt.Select(Filter, Sort & " ASC")
                    Else
                        SortedRows = dt.Select(Filter, Sort & " DESC")
                    End If
                End If
            End If

            NewDT = dt.Copy
            NewDT.Rows.Clear()
            For Row = 0 To SortedRows.GetUpperBound(0)
                Dim ItemArray(dt.Columns.Count - 1) As Object
                ItemArray = SortedRows(Row).ItemArray
                dr = NewDT.NewRow()
                NewDT.Rows.Add(dr)
                NewDT.Rows(Row).ItemArray = ItemArray
            Next

            Return NewDT

        Catch ex As Exception
            Throw New Exception("Error #CM2208: Common DTSort. " & ex.Message, ex)
        End Try
    End Function

    Public Function CloneColumn(ByVal SourceColumn As DataColumn) As DataColumn
        Dim ColumnName As String
        Dim ColumnType As Type
        Dim NewColumn As DataColumn

        ColumnName = SourceColumn.ColumnName
        ColumnType = SourceColumn.DataType
        NewColumn = New DataColumn(ColumnName, ColumnType)
        Return NewColumn
    End Function
#End Region
End Class