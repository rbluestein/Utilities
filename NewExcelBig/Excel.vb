Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
' Com -> Microsoft Excel 11.0 Object library

'' Start a new workbook in Excel.
'' ===========================
''oExcel = CreateObject("Excel.Application")
'Dim oExcel As New Microsoft.Office.Interop.Excel.Application
'            oBook = oExcel.Workbooks.Add
''oSheet = oBook.Worksheets(1)
'Dim oSheet As New Microsoft.Office.Interop.Excel.Worksheet
'            oSheet = oBook.Worksheets(1)
'            oSheet.Range("1:1").Font.Bold = True


' // Object Browser: View --> Object Browser
' // Microsoft.Office.Interop.Excel.Constants

Public Class Excel
    Public Function ExportToExcel(ByRef dt As DataTable, Optional ByVal SheetName As String = Nothing) As Results
        Dim i As Integer
        Dim oBook As Object
        Dim Row As Integer
        Dim Col As Integer
        Dim DataArray As Object
        Dim CharCount As Integer
        Dim DataType As String
        Dim Max As Integer = 911 '2000 '911
        Dim NumberDimensions As Integer
        Dim TooBig As New Collection
        Dim ExcelRow As Integer
        Dim ExcelCol As String
        Dim Results As New Results

        Try

            ' ___ Start a new workbook in Excel.
            Dim oExcel As New Microsoft.Office.Interop.Excel.Application
            oBook = oExcel.Workbooks.Add
            'oSheet = oBook.Worksheets(1)
            Dim oSheet As New Microsoft.Office.Interop.Excel.Worksheet
            oSheet = oBook.Worksheets(1)

            ' ___ Set zoom
            oExcel.ActiveWindow.Zoom = 85

            oSheet.Range("1:1").Font.Bold = True

            ReDim DataArray(dt.Rows.Count, dt.Columns.Count - 1)

            ' ___ Column headings
            For Col = 0 To dt.Columns.Count - 1
                DataArray(0, Col) = dt.Columns(Col).ColumnName
            Next
            ExcelRow = 2
            ExcelCol = "A"

            ' ___ Data
            For Row = 0 To dt.Rows.Count - 1
                For Col = 0 To dt.Columns.Count - 1
                    DataType = dt.Columns.Item(Col).DataType.ToString().ToUpper

                    If DataType = "SYSTEM.GUID" Then
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            'DataArray(Row + 1, Col) = dt.Rows(Row)(Col).ToString
                            DataArray(Row + 1, Col) = "'" & dt.Rows(Row)(Col).ToString
                        End If
                    ElseIf DataType = "SYSTEM.STRING" Then
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            CharCount = dt.Rows(Row)(Col).length
                            If CharCount > Max Then
                                Dim BigData(1) As Object
                                BigData(0) = ExcelCol & ExcelRow.ToString
                                'BigData(1) = dt.Rows(Row)(Col)
                                BigData(1) = "'" & dt.Rows(Row)(Col)
                                TooBig.Add(BigData)
                                DataArray(Row + 1, Col) = String.Empty
                            Else

                                'oSheet.Columns("B:B").Select()
                                'oSheet.Selection.NumberFormat = "@"
                                'DataArray(Row + 1, Col) = dt.Rows(Row)(Col)
                                DataArray(Row + 1, Col) = "'" & dt.Rows(Row)(Col)
                            End If
                        End If
                    Else
                        If IsDBNull(dt.Rows(Row)(Col)) Then
                            DataArray(Row + 1, Col) = String.Empty
                        Else
                            DataArray(Row + 1, Col) = dt.Rows(Row)(Col)
                        End If
                    End If
                    ExcelCol = AddColumn(ExcelCol)
                Next
                ExcelCol = "A"
                ExcelRow += 1
            Next
            If dt.Columns.Count = 1 Then
                NumberDimensions = 1
            End If

            If NumberDimensions = 1 Then
                Dim TempArray(DataArray.getupperbound(0), 1) As Object
                For i = 0 To TempArray.GetUpperBound(0)
                    TempArray(i, 0) = DataArray(i, 0)
                    TempArray(i, 1) = String.Empty
                Next
                DataArray = TempArray
            End If

            oSheet.Range("A1").Resize(DataArray.GetUpperBound(0) + 1, DataArray.GetUpperBound(1) + 1).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlTop

            oSheet.Range("A1").Resize(DataArray.GetUpperBound(0) + 1, DataArray.GetUpperBound(1) + 1).Value = DataArray
            If TooBig.Count > 0 Then
                For i = 1 To TooBig.Count
                    oSheet.Range(TooBig(i)(0)).Value = TooBig(i)(1)
                Next
            End If

            If SheetName <> Nothing Then
                oSheet.Name = SheetName
            End If

            oExcel.ActiveWindow.Zoom = 75
            oSheet.Columns.AutoFit()

            ' With oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, iNumCols))
            'With oSheet.Range("3:3")
            '    .Font.Bold = True
            '    .Font.Size = 11
            '    '   .Interior.Color = Microsoft.Office.Interop.Excel.Constants.xlColor1
            '    .Interior.Color = Microsoft.Office.Interop.Excel.Constants.xlColor1
            '    '  .Borders(c.xlEdgeBottom).Weight = c.xlThick
            'End With

            '   oSheet.Range("1:1").Select()
            'With oSheet.Range("4:4").Interior
            '    .ColorIndex = 6
            '    ' .Pattern = xlSolid
            'End With


            '  .Range("A1").BorderAround(, XlBorderWeight.xlThin, Microsoft.Office.Interop.XlColorIndex.xlColorIndexAutomatic)

            '  oSheet.Range("1:1").EntireRow.Borders(Excel.XlBordersIndex.xlEdgeBottom)


            '  oSheet.Range("1:1").Select()
            '    oSheet.Range("1:1").Interior.Color = excel.
            ' oSheet.Range("1:1").Pattern = Excel.xlSolid

            '  End With


            'Rows("1:1").Select()
            'With Selection.Interior
            '    .ColorIndex = 6
            '    .Pattern = xlSolid
            'End With



            DataArray = Nothing

            oExcel.Visible = True
            oExcel.DisplayAlerts = False
            oSheet = Nothing
            oBook = Nothing
            'oExcel.Quit()
            oExcel = Nothing
            GC.Collect()

            Results.Success = True
            Return Results

        Catch ex As Exception
            Results.Success = False
            Results.Msg = ex.Message
            Return Results
        End Try
    End Function

    Public Function AddColumn(ByVal ColName As String) As String
        Dim i As Integer
        Dim MaxLength As Integer = 5
        Dim Rec_(MaxLength - 1) As String
        Dim Results As String = Nothing
        Dim Carry As Boolean

        Try

            ' Pad ColName
            ColName = ColName.PadLeft(MaxLength)

            ' Load the column value into Rec_
            For i = 0 To MaxLength - 1
                Rec_(i) = ColName.Substring(i, 1)
            Next

            ' Begin processing
            For i = MaxLength - 1 To 0 Step -1
                If i = MaxLength - 1 Then
                    ' Perform alpha addition on the rightmost character.Apply carry if applicable.
                    If Rec_(i) = "Z" Then
                        Rec_(i) = "A"
                        Carry = True
                    Else
                        Rec_(i) = Chr(Asc(Rec_(i)) + 1)
                    End If
                Else   ' Not the rightmost character

                    ' If the character to its right resulted in a carry...
                    If Carry Then
                        If Asc(Rec_(i)) = 32 Then
                            Rec_(i) = "A"
                            Carry = False
                        ElseIf Rec_(i) = "Z" Then
                            Rec_(i) = "A"
                            Carry = True
                        Else
                            Rec_(i) = Chr(Asc(Rec_(i)) + 1)
                            Carry = False
                        End If

                    End If
                End If
            Next

            For i = 0 To Rec_.GetUpperBound(0)
                Results &= Rec_(i)
            Next
            Return Trim(Results)

        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #504: Excel.AddColumn " & ex.Message)
        End Try
    End Function


    Public Function ExcelToDT(ByVal FullPath As String, Optional ByVal Sql As String = "") As DataTable
        'Dim ConnStr As String
        'ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Inetpub\wwwroot\ImportExportManagement\MasterList.xls;Extended Properties=Excel 8.0;"

        Try
            Dim dt As New DataTable
            Dim Header As Boolean

            If Sql.Length = 0 Then
                Sql = "SELECT * FROM [Sheet1$]"
            End If

            'Dim da As New OleDbDataAdapter("SELECT * FROM [Feed_DTS$]", ConnStr)
            '   Dim da As New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [ApptLicFieldsOnly$]", GetExcelConnectionString("License.xls", True))
            ' Dim da As New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & WorksheetName & "$]", GetExcelConnectionString(Filename, Header))
            Dim da As New System.Data.OleDb.OleDbDataAdapter(Sql, GetExcelConnectionString(FullPath, Header))
            da.Fill(dt)
            Return dt

        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #503: Excel.ExcelToDT " & ex.Message)
        End Try
    End Function


    Private Function GetExcelConnectionString(ByVal FullPath As String, ByVal Header As Boolean) As String
        ' http://www.connectionstrings.com/?carrier=excel

        Try

            Dim ConnString As String
            ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<fullpath>;Extended Properties=""Excel 8.0;HDR=<header>"";"

            If System.Environment.MachineName.ToUpper = "LT-5ZFYRC1" Then
                ConnString = Replace(ConnString, "<fullpath>", FullPath)
                If Header Then
                    ConnString = Replace(ConnString, "<header>", "Yes")
                Else
                    ConnString = Replace(ConnString, "<header>", "No")
                End If
            End If
            Return ConnString

        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #504: Excel.GetExcelConnectionString " & ex.Message)
        End Try
    End Function

    Public Function ExcelToTbl(ByVal FullPath As String, Optional ByVal Sql As String = "") As DataTable
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable
        Dim ConnString As String

        Try

            If Sql.Length = 0 Then
                Sql = "SELECT * FROM [Sheet1$]"
            End If
            'ConnString = "provider=Microsoft.Jet.OLEDB.4.0; " & "data source='" & FullPath & " '; " & "Extended Properties=Excel 8.0;"
            ConnString = "data source='" & FullPath & " '; " & "Extended Properties=Excel 8.0;"


            '"user id=BVI_SQL_SERVER;password=noisivtifeneb;database=|;server="

            Dim SqlCmd As New SqlCommand(Sql)
            SqlCmd.CommandType = CommandType.Text
            SqlCmd.Connection = New SqlConnection(ConnString)
            DataAdapter = New SqlDataAdapter(SqlCmd)
            DataAdapter.Fill(dt)
            DataAdapter.Dispose()
            SqlCmd.Dispose()
            Return dt

        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #505: Excel.ExcelToTbl " & ex.Message)
        End Try
    End Function


    'Public Function ExcelToTbl()
    '    Dim MyConnection As System.Data.OleDb.OleDbConnection
    '    Dim FullPath As String

    '    Try
    '        FullPath = "C:\Apps\UserManagement\MigrationOutput.xls"
    '        ''''''' Fetch Data from Excel
    '        Dim DtSet As System.Data.DataSet
    '        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
    '        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & "data source='" & FullPath & " '; " & "Extended Properties=Excel 8.0;")

    '        ' Select the data from Sheet1 of the workbook.
    '        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
    '        MyCommand.TableMappings.Add("Table", "Attendence")
    '        DtSet = New System.Data.DataSet
    '        MyCommand.Fill(DtSet)

    '        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '        'DataGrid1.DataSource = DtSet.Tables(0)
    '        MyConnection.Close()

    '    Catch ex As Exception
    '        MyConnection.Close()
    '    End Try
    'End Function

    ' *  Connection String
    '       Syntax: Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<Full Path of Excel File>; Extended Properties="Excel 8.0; HDR=No; IMEX=1".

    'Definition of Extended Properties:
    '    * Excel = <No>
    '      One should specify the version of Excel Sheet here. For Excel 2000 and above, it is set it to Excel 8.0 and for all others, it is Excel 5.0.

    '    * HDR= <Yes/No>
    '      This property will be used to specify the definition of header for each column. If the value is ‘Yes’, the first row will be treated as heading. Otherwise, the heading will be generated by the system like F1, F2 and so on.

    '    * IMEX= <0/1/2>
    '      IMEX refers to IMport EXport mode. This can take three possible values.
    '          o IMEX=0 and IMEX=2 will result in ImportMixedTypes being ignored and the default value of ‘Majority Types’ is used. In this case, it will take the first 8 rows and then the data type for each column will be decided.
    '          o IMEX=1 is the only way to set the value of ImportMixedTypes as Text. Here, everything will be treated as text. 

    'For more info regarding Extended Properties, http://www.dicks-blog.com/archives/2004/06/03/
End Class


