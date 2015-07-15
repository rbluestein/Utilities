Imports System.Data.SqlClient
Imports System.Data

Public Class DBase
    Public Function SqlProperCase(ByVal SubjectFieldName As String, ByVal NewFieldName As String) As String
        Dim Results As String = String.Empty
        If SubjectFieldName.Length = 1 Then
            Results = Replace("UPPER(SUBSTRING(|,1, 1)) AS ", "|", SubjectFieldName)
            Results &= NewFieldName
        Else
            Results = Replace("UPPER(SUBSTRING(|,1, 1))  + LOWER(SUBSTRING(|, 2, LEN(|) -1)) AS ", "|", SubjectFieldName)
            Results &= NewFieldName
        End If
        Return Results
    End Function

    Public Class QueryPack
        Private cReturnDataTable As Boolean
        Private cReturnDataSet As Boolean
        Private cSuccess As Boolean
        Private cGenErrMsg As String
        Private cTechErrMsg As String
        Private cdt As DataTable
        Private cds As DataSet

        Public Property Success()
            Get
                Return cSuccess
            End Get
            Set(ByVal Value)
                cSuccess = Value
            End Set
        End Property

        Public ReadOnly Property GenErrMsg()
            Get
                Return GenErrMsg
            End Get
        End Property
        Public Property TechErrMsg()
            Get
                Return cTechErrMsg
            End Get
            Set(ByVal Value)
                cTechErrMsg = Value
            End Set
        End Property
        Public Property dt()
            Get
                Return cdt
            End Get
            Set(ByVal Value)
                cdt = Value
            End Set
        End Property
        Public Property ds()
            Get
                Return cds
            End Get
            Set(ByVal Value)
                cds = Value
            End Set
        End Property
    End Class

    Public Class CmdPack
        Private cSqlCmd As SqlClient.SqlCommand
        Private cParameterColl As New Collection
        Private cTechErrMsg As String
        Private cSuccess As Boolean

        Public Sub New(ByVal SPNameOrSqlText As String, ByVal CmdType As CommandType, ByVal Enviro As Enviro)
            cSqlCmd = New SqlClient.SqlCommand(SPNameOrSqlText)
            cSqlCmd.CommandType = CmdType
            cSqlCmd.Connection = New SqlConnection(Enviro.GetConnectionString(Enviro.DBHost, Enviro.DefaultDatabase))
        End Sub

        Public Sub AddParameter(ByVal ParameterType As SqlDbType, ByVal VarName As String, ByVal Value As Object, ByVal Direction As ParameterDirection, Optional ByVal Length As Integer = 0)
            Dim Parameter As SqlParameter
            Parameter = New SqlParameter(VarName, SqlDbType.VarChar, Length)
            Parameter.Value = Value
            Parameter.Direction = Direction
            cSqlCmd.Parameters.Add(Parameter)
            If ParameterType = SqlDbType.VarChar Then
                Parameter.Size = Length
            End If
        End Sub

        Public Sub ExecuteNonQuery()
            Dim DataReader As SqlDataReader

            Try
                cSqlCmd.Connection.Open()
                DataReader = cSqlCmd.ExecuteReader()
                Dim Param As SqlParameter
                For Each Param In cSqlCmd.Parameters
                    cParameterColl.Add(Param.Value, Param.ParameterName)
                Next
                cSuccess = True

            Catch ex As Exception
                cSuccess = False
                cTechErrMsg = ex.Message
            Finally
                Try
                    DataReader.Close()
                Catch
                End Try
                cSqlCmd.Dispose()
                cSqlCmd.Connection.Close()
            End Try
        End Sub

        Public ReadOnly Property Success() As Boolean
            Get
                Return cSuccess
            End Get
        End Property
        Public ReadOnly Property TechErrMsg() As String
            Get
                Return cTechErrMsg
            End Get
        End Property
        Public ReadOnly Property ParameterColl() As Collection
            Get
                Return cParameterColl
            End Get
        End Property
    End Class

#Region " Formatted table "
    Public Function GetDTExtended(ByRef dt As DataTable) As DataTable
        Dim Row As Integer
        Dim Col As Integer
        Dim dt2 As New DataTable
        Dim dr2 As DataRow

        Try

            For Col = 0 To dt.Columns.Count - 1
                dt2.Columns.Add(dt.Columns(Col).ColumnName, GetType(ExtendedCell))
            Next

            For Row = 0 To dt.Rows.Count - 1
                dr2 = dt2.NewRow
                For Col = 0 To dt.Columns.Count - 1
                    Select Case dt.Columns(Col).DataType.ToString.ToLower
                        Case "system.string", "system.guid"
                            Dim ExtendedCellVarchar As New ExtendedCellVarchar(dt.Rows(Row)(Col))
                            Dim ExtendedCell As ExtendedCell = ExtendedCellVarchar
                            dr2(Col) = ExtendedCell
                        Case "system.int64", "system.int32", "system.int16"
                            Dim ExtendedCellInt As New ExtendedCellInt(dt.Rows(Row)(Col))
                            Dim ExtendedCell As ExtendedCell = ExtendedCellInt
                            dr2(Col) = ExtendedCell
                        Case "system.boolean"
                            Dim ExtendedCellBit As New ExtendedCellBit(dt.Rows(Row)(Col))
                            Dim ExtendedCell As ExtendedCell = ExtendedCellBit
                            dr2(Col) = ExtendedCell
                        Case "system.datetime"
                            Dim ExtendedCellDateTime As New ExtendedCellDateTime(dt.Rows(Row)(Col))
                            Dim ExtendedCell As ExtendedCell = ExtendedCellDateTime
                            dr2(Col) = ExtendedCell
                        Case "system.double", "system.single"
                            Dim ExtendedCellFloat As New ExtendedCellFloat(dt.Rows(Row)(Col))
                            Dim ExtendedCell As ExtendedCell = ExtendedCellFloat
                            dr2(Col) = ExtendedCell
                        Case "system.decimal"
                            Dim ExtendedCellMoney As New ExtendedCellMoney(dt.Rows(Row)(Col))
                            Dim ExtendedCell As ExtendedCell = ExtendedCellMoney
                            dr2(Col) = ExtendedCell
                        Case Else
                            Throw New Exception("Unprocessed data type in DBase.GetDTExtended -- " & dt.Columns(Col).DataType.ToString.ToLower & ".")
                    End Select
                Next
                dt2.Rows.Add(dr2)
            Next
            Return dt2

        Catch ex As Exception
            Throw New Exception("Error #2702: DBase GetDTExtended. " & ex.Message, ex)
        End Try
    End Function

    Public Class ExtendedCell
        Protected cValue As Object
        Public Sub New(ByVal Value As Object)
            cValue = Value
        End Sub
    End Class

    Public Class ExtendedCellDateTime
        Inherits ExtendedCell

        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, DateTime)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDate(Value) Then
                    cValue = CType(Value, DateTime)
                ElseIf IsDBNull(Value) Then
                    cValue = DBNull.Value
                ElseIf Value = Nothing Then
                    cValue = DBNull.Value
                Else
                    cValue = CType(Value, DateTime)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    If cValue = "1/1/1950" Then
                        Return String.Empty
                    ElseIf cvalue = "1/1/1900" Then
                        Return String.Empty
                    Else
                        Return CType(cValue, DateTime)
                    End If
                    'Return CType(cValue, DateTime)
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                ' Return ToText
                Return CType(ToText, System.String)
            End Get
        End Property
        Public Function ToFormat(ByVal FormatString As String)
            If IsDate(ToText) Then
                Return CType(ToText, System.DateTime).ToString(FormatString)
            Else
                Return String.Empty
            End If
        End Function
    End Class

    Public Class ExtendedCellInt
        Inherits ExtendedCell
        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, System.Int64)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) Then
                    cValue = DBNull.Value
                ElseIf Value = Nothing Then
                    cValue = DBNull.Value
                Else
                    cValue = CType(Value, System.Int64)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    Return CType(cValue, System.Int64)
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class

    Public Class ExtendedCellVarchar
        Inherits ExtendedCell
        Private cLength As Integer

        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Length() As Integer
            Get
                Return cLength
            End Get
            Set(ByVal Value As Integer)
                cLength = Value
            End Set
        End Property
        Public Property Value()
            Get
                If IsDBNull(cValue) OrElse cValue = Nothing Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, System.String)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) OrElse Value = Nothing Then
                    cValue = DBNull.Value
                ElseIf IsDate(Value) Then
                    cValue = CType(Value, System.String)
                ElseIf IsNumeric(Value) Then
                    cValue = CType(Value, System.String)
                ElseIf Value = "null" Then
                    cValue = "null"
                Else
                    cValue = CType(Value, System.String)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                Dim Working As String
                If IsDBNull(cValue) OrElse cValue = Nothing Then
                    Return String.Empty
                Else
                    Working = CType(cValue, System.String)
                    ' Return Replace(Working, "~", "'")
                    Return Working
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Dim Working As String
                If ToText.length > 0 Then
                    Working = ToText
                    'Return Replace(Working, "'", "&#39;")
                    Return Replace(Working, "'", "~")
                    'Return Replace(Working, "'", "&apos;")
                    Return ToText
                Else
                    Return ToText
                End If
            End Get
        End Property
    End Class

    Public Class ExtendedCellBit
        Inherits ExtendedCell

        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    If cValue Then
                        Return 1
                    Else
                        Return 0
                    End If
                    'Return CType(cValue, System.Int64)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) Then
                    cValue = DBNull.Value
                Else
                    If Value Then
                        cValue = 1
                    Else
                        cValue = 0
                    End If
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    If cValue Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            End Get
        End Property
        Public ReadOnly Property ToDropdown2State() As Integer
            Get
                If IsDBNull(cValue) Then
                    Return 0
                Else
                    If cValue Then
                        Return 1
                    Else
                        Return 0
                    End If
                End If
            End Get
        End Property
        Public ReadOnly Property ToDropdown3State() As Object
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    If cValue Then
                        Return 1
                    Else
                        Return 0
                    End If
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class

    Public Class ExtendedCellMoney
        Inherits ExtendedCell

        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    'Return CType(cValue, System.Decimal)
                    'Return FormatNumber(CType(cValue, System.Decimal), 2)
                    Return cValue
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) Then
                    cValue = DBNull.Value
                Else
                    cValue = CType(Value, System.Decimal)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    'Return CType(cValue, System.Decimal)
                    Return FormatNumber(CType(cValue, System.Decimal), 2)
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class

    Public Class ExtendedCellFloat
        Inherits ExtendedCell
        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, System.Double)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) Then
                    cValue = DBNull.Value
                Else
                    cValue = CType(Value, System.Double)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    Return CType(cValue, System.Double)
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class
#End Region

End Class

