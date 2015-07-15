Imports System.Data.SqlClient

Public Class Common
    Private Shared cDBHost As String
    Private Shared cDBName As String
    Private Shared cTblName As String
    Private Shared cConnStringTemplate As String = "user id=BVI_SQL_SERVER;password=noisivtifeneb;database=|;server="

    Public Shared Property DBHost() As String
        Get
            Return cDBHost
        End Get
        Set(ByVal Value As String)
            cDBHost = Value
        End Set
    End Property

    Public Shared Property DBName() As String
        Get
            Return cDBName
        End Get
        Set(ByVal value As String)
            cDBName = value
        End Set
    End Property
    Public Shared Property TblName() As String
        Get
            Return cTblName
        End Get
        Set(ByVal value As String)
            cTblName = value
        End Set
    End Property

    Public Shared Function GetDT(ByVal Sql As String) As DataTable
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable

        Try

            Dim SqlCmd As New SqlCommand(Sql)
            SqlCmd.CommandType = CommandType.Text
            SqlCmd.Connection = New SqlConnection(GetConnectionString)
            DataAdapter = New SqlDataAdapter(SqlCmd)
            DataAdapter.Fill(dt)
            DataAdapter.Dispose()
            SqlCmd.Dispose()
            Return dt
        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #402: Common.GetDT " & ex.Message)
        End Try
    End Function

    Public Shared Function GetDTWithQueryPack(ByVal Sql As String) As QueryPack
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable
        Dim QueryPack As QueryPack

        Try

            QueryPack = New QueryPack
            Dim SqlCmd As New SqlCommand(Sql)
            SqlCmd.CommandType = CommandType.Text
            SqlCmd.Connection = New SqlConnection(GetConnectionString)
            DataAdapter = New SqlDataAdapter(SqlCmd)
            Try
                DataAdapter.Fill(dt)
                QueryPack.Success = True
                QueryPack.dt = dt
            Catch ex As Exception
                QueryPack.Success = False
                QueryPack.TechErrMsg = ex.Message
            End Try

            DataAdapter.Dispose()
            SqlCmd.Dispose()
            Return QueryPack

        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #403: Common.GetDTWithQueryPack " & ex.Message)
        End Try
    End Function

    'Public Shared Function SqlServerConnectionString(ByVal DBName As String) As String
    '    Return Replace(cConnStringTemplate, "|", DBName) & Common.IPAddress
    'End Function

    Public Shared Function GetConnectionString() As String
        'Return ConnectionString(Common.IPAddress, "UserManagement")
        Return Replace(cConnStringTemplate, "|", DBName) & DBHost
    End Function

    'Public Shared Function ConnectionString(ByVal DBName As String) As String
    '    Return ConnectionString(Common.IPAddress, DBName)
    'End Function

    'Public Shared Function ConnectionString(ByVal DBHost As String, ByVal DBName As String) As String
    '    Return Replace(cConnStringTemplate, "|", DBName) & DBHost
    'End Function

    Public Shared Sub ExitApplication()
        Environment.Exit(0)
    End Sub
End Class

Public Class QueryPack
    Private cReturnDataTable As Boolean
    Private cReturnDataSet As Boolean
    Private cSuccess As Boolean
    Private cGenErrMsg As String
    Private cTechErrMsg As String
    Private cdt As DataTable
    Private cds As DataSet

    Public Property Success() As Boolean
        Get
            Return cSuccess
        End Get
        Set(ByVal Value As Boolean)
            cSuccess = Value
        End Set
    End Property

    Public ReadOnly Property GenErrMsg() As String
        Get
            Return cGenErrMsg
        End Get
    End Property
    Public Property TechErrMsg() As String
        Get
            Return cTechErrMsg
        End Get
        Set(ByVal Value As String)
            cTechErrMsg = Value
        End Set
    End Property
    Public Property dt() As DataTable
        Get
            Return cdt
        End Get
        Set(ByVal Value As DataTable)
            cdt = Value
        End Set
    End Property
    Public Property ds() As DataSet
        Get
            Return cds
        End Get
        Set(ByVal Value As DataSet)
            cds = Value
        End Set
    End Property
End Class

Public Class Results
    Public Success As Boolean
    Public Msg As String
    Public ResponseAction As ResponseAction
    Public ObtainConfirm As Boolean
End Class

Public Enum ResponseAction
    DisplayBlank = 1
    DisplayUserInputNew = 2
    DisplayUserInputExisting = 3
    DisplayExisting = 4
    ReturnToCallingPage = 5
End Enum

Public Class Report
    Public Enum ReportTypeEnum
        Information = 1
        InformationNoLog = 2
        [Error] = 3
        ErrorNoShutdown = 4
        MessageBoxAndShutdown = 5
    End Enum

    Public Shared Sub Report(ByVal ReportType As ReportTypeEnum, ByVal Message As String, Optional ByVal Caption As String = Nothing)
        MessageBox.Show(Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Common.ExitApplication()
    End Sub
End Class