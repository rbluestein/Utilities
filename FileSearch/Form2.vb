Imports System.IO

Public Class Form2
    Private cCancel As Boolean
    'Private cWarningColl As New CollX

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            ' GenerateError()

            lstResults.HorizontalScrollbar = True

            Me.Text = "World Famous File Search  v" & GetVersionNumber()

            ' ___ If launched from the Send To menu, use the argument as our directory
            If My.Application.CommandLineArgs.Count > 0 Then
                txtSearchPath.Text = My.Application.CommandLineArgs(0)
                txtSearchFor.Focus()
            End If

        Catch ex As Exception
            DisplayError("Error #1002: FileSearch.Form_Load. ", ex)
        End Try
    End Sub

    Private Function GetVersionNumber() As String
        Dim VersionNumber As String
        'VersionNumber = "1.01.00"  '12/07/2013: Added error handling for invalid path. Fixed tab order. Show version in form heading.
        VersionNumber = "1.01.01"  '12/13/2013: Added comprehensive error handling.
        Return VersionNumber
    End Function

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim Dir_Info As DirectoryInfo
        Dim DirColl As New CollX
        Dim FileColl As New CollX
        Dim Idx As Integer

        Try

            Dir_Info = New DirectoryInfo(txtSearchPath.Text.Trim)
            If Not Dir_Info.Exists Then
                ShowProgress("Search cancelled")
                DisplayError("Path """ & txtSearchPath.Text.Trim & """ does not exist", DBNull.Value)
                'MessageBox.Show("Path """ & txtSearchPath.Text.Trim & """ does not exist", "Error", MessageBoxButtons.OK)
                Exit Sub
            End If

            cCancel = False
            lstResults.Items.Clear()

            DirColl.Assign(txtSearchPath.Text)
            Idx = 1

            Do
                AddSubDirs(DirColl, Idx)

                If cCancel Then
                    Exit Do
                End If

                Idx += 1
                If Idx > DirColl.Count Then
                    Exit Do
                End If
            Loop

            If Not cCancel Then
                For Idx = 1 To DirColl.Count
                    If cCancel Then
                        Exit For
                    End If
                    ProcessSearch(DirColl, Idx)
                Next
            End If

            If cCancel Then
                ShowProgress("Search cancelled")
            Else
                ShowProgress("Search complete")
            End If

        Catch ex As Exception
            DisplayError("Error #1003: FileSearch.btnSearchClick. ", ex)
        End Try
    End Sub

    Private Sub ProcessSearch(ByRef DirColl As CollX, ByVal Idx As Integer)
        Dim FileCompleteText As String
        Dim Dir_Info As DirectoryInfo
        Dim _File_Info() As FileInfo = Nothing
        Dim FileNameOrPattern As String
        Dim MatchString As String
        Dim ErrorNum As Integer

        Try

            If cboFileOrPattern.Text.Trim.Length = 0 Then
                FileNameOrPattern = "*.*"
            Else
                FileNameOrPattern = cboFileOrPattern.Text.Trim
            End If

            MatchString = txtSearchFor.Text

            ' ___  Get the files in this directory
            Dir_Info = New DirectoryInfo(DirColl(Idx))
            Try
                _File_Info = Dir_Info.GetFiles(FileNameOrPattern)
            Catch
                ErrorNum = 1
            End Try

            ' ___ No content specified: display full path. Otherwise, search against content.
            If ErrorNum = 0 Then
                For Each File_Info As FileInfo In _File_Info
                    If MatchString.Length = 0 Then
                        lstResults.Items.Add(File_Info.FullName)
                        System.Windows.Forms.Application.DoEvents()
                    Else
                        FileCompleteText = File.ReadAllText(File_Info.FullName)
                        If FileCompleteText.IndexOf(MatchString, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            lstResults.Items.Add(File_Info.FullName)
                            System.Windows.Forms.Application.DoEvents()
                        End If
                    End If
                Next File_Info
                _File_Info = Nothing
            End If

        Catch ex As Exception
            DisplayError("Error #1004: FileSearch.ProcessSearch. ", ex)
        End Try
    End Sub

    Private Sub xAddSubDirs(ByRef DirColl As CollX, ByVal Idx As Integer)
        Dim Dir_Info As DirectoryInfo
        Dim _SubDir() As DirectoryInfo

        Try
            Dir_Info = New DirectoryInfo(DirColl(Idx))
            _SubDir = Dir_Info.GetDirectories()
            For Each SubDir As DirectoryInfo In _SubDir
                DirColl.Assign(SubDir.FullName)
                ShowProgress(SubDir.FullName)
            Next SubDir
        Catch ex As Exception
            DisplayError("Error #1005: FileSearch.AddSudDirs. ", ex)
        End Try
    End Sub

    Private Sub AddSubDirs(ByRef DirColl As CollX, ByVal Idx As Integer)
        Dim Dir_Info As DirectoryInfo
        Dim _SubDir() As DirectoryInfo = Nothing
        Dim ErrorNum As Integer

        Try
            Dir_Info = New DirectoryInfo(DirColl(Idx))
            Try
                _SubDir = Dir_Info.GetDirectories()
            Catch ex As Exception
                ErrorNum = 1
                'cWarningColl.Assign(ex.Message)
                'UserChoice = MessageBox.Show(ex.Message & " Continue?", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
                'If UserChoice = 2 Then
                '    cCancel = True
                'End If
            End Try

            If ErrorNum = 0 Then
                For Each SubDir As DirectoryInfo In _SubDir
                    DirColl.Assign(SubDir.FullName)
                    ShowProgress(SubDir.FullName)
                Next SubDir
            End If
        Catch ex As Exception
            DisplayError("Error #1005: FileSearch.AddSudDirs. ", ex)
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        'Shell("NewExcelBig.exe prod", AppWinStyle.NormalFocus)
        Try
            Shell("FileSearch.exe", AppWinStyle.NormalFocus)
        Catch ex As Exception
            DisplayError("Error #1006: FileSearch.btnNewClick. ", ex)
        End Try
    End Sub

    Private Sub ShowProgress(ByVal Status As String)
        Try
            txtProgress.Text = Status
            System.Windows.Forms.Application.DoEvents()
        Catch ex As Exception
            DisplayError("Error #1007: FileSearch.ShowProgress. ", ex)
        End Try
    End Sub

    Private Sub GenerateError()
        Dim a, b, c As Integer
        c = b / a
    End Sub

    Private Sub DisplayError(ByVal Prefix As String, ByVal Input As Object)
        Dim Message As String

        Try

            If IsDBNull(Input) Then
                Message = Prefix
            ElseIf TypeOf (Input) Is Exception Then
                Message = Prefix & Input.message
            Else
                Message = Prefix & Input
            End If

            MessageBox.Show(Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cCancel = True
            ' Application.Exit()
        Catch ex As Exception


        End Try
    End Sub


    'Private Sub ProcessSearchRecursive(ByVal lstResults As ListBox, ByVal FileName As String, ByVal Dir_Info As DirectoryInfo, ByVal Content As String)
    '    Dim FileCompleteText As String
    '    Dim _File_Info() As FileInfo
    '    Dim _SubDir() As DirectoryInfo

    '    ' ___  Get the files in this directory
    '    ' Dim fs_infos() As FileInfo = dir_info.GetFiles()
    '    _File_Info = Dir_Info.GetFiles(FileName)

    '    ' ___ No content specified: display full path. Otherwise, search against content.
    '    For Each File_Info As FileInfo In _File_Info
    '        If Content.Length = 0 Then
    '            lstResults.Items.Add(File_Info.FullName)
    '            ShowProgress(File_Info.FullName)
    '        Else
    '            FileCompleteText = File.ReadAllText(File_Info.FullName)
    '            If FileCompleteText.IndexOf(Content, StringComparison.OrdinalIgnoreCase) >= 0 Then
    '                lstResults.Items.Add(File_Info.FullName)
    '                ShowProgress(File_Info.FullName)
    '            End If
    '        End If
    '    Next File_Info
    '    _File_Info = Nothing

    '    _SubDir = Dir_Info.GetDirectories()
    '    For Each SubDir As DirectoryInfo In _SubDir
    '        ShowProgress(SubDir.FullName)
    '        ProcessSearch(lstResults, FileName, SubDir, Content)
    '    Next SubDir
    'End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        cCancel = True
    End Sub
End Class


Public Class CollX
    Inherits System.Collections.CollectionBase
    ' Private Bittem As ListItem

    Public Sub New()
        List.Add(DBNull.Value)
    End Sub

    Public Overloads ReadOnly Property Count() As Integer
        Get
            Return List.Count - 1
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Idx As Integer) As Object
        Get
            Return List(Idx).Value
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Key As String) As Object
        Get
            Dim i As Integer
            Dim KeyUpper As String
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            Throw New CollXError("Error #3604: CallX.Coll item not found error. Key: " & Key)  'CollXError
        End Get
    End Property

    Public Function TreatKeyAsString(ByVal Key As String) As Object
        Dim i As Integer
        Dim KeyUpper As String

        Try
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            Throw New CollXError("Error #3611: CallX.GetValueKeyAsString item not found error. Key: " & Key)  'CollXError

        Catch ex As Exception
            Throw New CollXError("Error #3611: CallX.GetValueKeyAsString item not found error. Key: " & Key)  'CollXError
        End Try
    End Function
    Public ReadOnly Property Key(ByVal Idx As Integer) As String
        Get
            Dim i As Integer
            For i = 1 To List.Count - 1
                If i = Idx Then
                    'Return List(i).Value
                    Return List(i).Key
                End If
            Next
            Return Nothing
        End Get
    End Property
    Public ReadOnly Property DoesKeyExist(ByVal Key As String) As Boolean
        Get
            Dim i As Integer
            Key = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = Key Then
                    Return True
                End If
            Next
            Return False
        End Get
    End Property

    Public Sub Assign(ByVal Key As String, ByVal Value As Object)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key Then
                    List(i).Value = Nothing
                    List(i).Value = Value
                    Found = True
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                List.Add(New KeyValuePair(Key, Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3604: CallX.Assign. Item not found error. Key: " & Key)
        End Try
    End Sub

    Public Sub Assign(ByVal Key_Value As String)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key_Value Then
                    List(i).Value = Nothing
                    List(i).Value = Key_Value
                    Found = True
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                List.Add(New KeyValuePair(Key_Value, Key_Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3605: CallX.Assign. " & ex.Message)
        End Try
    End Sub

    'Public Sub ConvertArr(ByRef obj As Object)
    '    Try

    '        Dim i As Integer
    '        For i = 0 To obj.GetUpperBound(0)
    '            Assign(obj(i))
    '        Next

    '    Catch ex As Exception
    '        Throw New CollXError("Error #3606: CallX.ConvertArr. " & ex.Message)
    '    End Try
    'End Sub

    Public Sub ConvertRow(ByRef dr As DataRow)
        Try

            Dim i As Integer
            For i = 0 To dr.ItemArray.GetUpperBound(0)
                Assign(dr.Table.Columns(i).ColumnName, dr(i))
            Next

        Catch ex As Exception
            Throw New CollXError("Error #3610: CallX.ConvertRow. " & ex.Message)
        End Try
    End Sub

    Public Function ConvertToStr(ByVal Delimiter As String) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        Try

            For i = 1 To List.Count - 1
                If i < List.Count - 1 Then
                    sb.Append(List(i).Value & Delimiter)
                Else
                    sb.Append(List(i).Value)
                End If
            Next
            Return sb.ToString
        Catch ex As Exception
            Throw New CollXError("Error #3612: CallX.ConvertToStr. " & ex.Message)
        End Try
    End Function

    'Public Sub ConvertStr(ByRef Input As String, ByRef Delimiter As String)
    '    Dim i As Integer
    '    Dim Box As String()

    '    Try

    '        If Input.Length > 0 Then
    '            If Input.Substring(Input.Length - 1) = Delimiter Then
    '                Input = Input.Substring(0, Input.Length - 1)
    '            End If
    '            Box = Split(Input, Delimiter)
    '            For i = 0 To Box.GetUpperBound(0)
    '                Assign(Box(i))
    '            Next
    '        End If

    '    Catch ex As Exception
    '        Throw New CollXError("Error #3607: CallX.ConvertStr. " & ex.Message)
    '    End Try
    'End Sub

    Public Function CollxToSql() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        For i = 1 To List.Count - 1
            If i < List.Count - 1 Then
                sb.Append(List(i).Key & "=" & List(i).Value & ", ")
            Else
                sb.Append(List(i).Key & "=" & List(i).Value)
            End If
        Next
        Return sb.ToString
    End Function

#Region " New from... "
    Public Shared Function NewFromList(ByVal Input As String, ByVal Delimiter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Coll As New CollX
        Box = Input.Split("|")
        If Box.GetUpperBound(0) > -1 Then
            For i = 0 To Box.GetUpperBound(0)
                Coll.Assign(Box(i))
            Next
        End If
        Return Coll
    End Function

    Public Shared Function NewFromDataRow(ByRef dr As DataRow) As CollX
        Dim i As Integer
        Dim dt As DataTable
        Dim Coll As New CollX
        dt = dr.Table
        For i = 0 To dt.Columns.Count - 1
            Coll.Assign(dt.Columns(i).ColumnName, dr(i))
        Next
        Return Coll
    End Function

    Public Shared Function NewFromTable(ByRef dt As DataTable) As CollX
        Dim i As Integer
        Dim Coll As New CollX

        If dt.Columns.Count = 1 Then
            For i = 0 To dt.Rows.Count - 1
                Coll.Assign(dt.Rows(i)(0), dt.Rows(i)(0))
            Next
        Else
            For i = 0 To dt.Rows.Count - 1
                Coll.Assign(dt.Rows(i)(0), dt.Rows(i)(1))
            Next
        End If

        Return Coll
    End Function


    Public Shared Function NewFromKeyValue(ByVal Input As String, ByVal RowDelimter As String, ByVal ColDelimter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Box2 As String()
        Dim Coll As New CollX
        Box = Input.Split(RowDelimter)
        For i = 0 To Box.GetUpperBound(0)
            Box2 = Box(i).Split(ColDelimter)
            Coll.Assign(Box2(0), Box2(1))
        Next
        Return Coll
    End Function
#End Region

    Public Overloads Sub RemoveAt(ByVal Index As Integer)
        List.RemoveAt(Index)
    End Sub

    Public Overloads Sub Remove(ByVal Key As String)
        Dim i As Integer
        For i = 1 To List.Count - 1
            If List(i).Key.ToUpper = Key.ToUpper Then
                List.Remove(List(i))
                Exit For
            End If
        Next
    End Sub

    Public Shared Function GetDistinct(ByRef dt As DataTable, ByVal FieldName As String) As CollX
        Dim i As Integer
        Dim Coll As New CollX
        Dim FirstValue As Object = DBNull.Value

        If dt.Rows.Count = 0 Then
            Return Nothing
        Else
            Coll.Assign(dt.Rows(0)(FieldName))
            For i = 0 To dt.Rows.Count - 1
                If dt.Rows(0)(FieldName) <> Coll(Coll.Count) Then
                    Coll.Assign(dt.Rows(i)(FieldName))
                End If
            Next
        End If
        Return Coll
    End Function

    Public Function View() As String()
        Dim Output(Me.Count) As String
        Dim Val As String

        Try

            For i = 1 To List.Count - 1
                Try
                    Val = List(i).Value
                Catch ex As Exception
                    Val = "<object>"
                End Try
                Output(i) = List(i).Key & "|" & Val
            Next
            Return Output


        Catch ex As Exception
            Throw New CollXError("Error #3608: CallX.View. " & ex.Message)
        End Try
    End Function

    Public Shared Function Clone(ByVal InputColl As CollX) As CollX
        Dim i As Integer
        Dim OutputColl As New CollX

        Try

            For i = 1 To InputColl.Count
                OutputColl.Assign(InputColl.Key(i), InputColl(i))
            Next
            Return OutputColl

        Catch ex As Exception
            Throw New CollXError("Error #3609: CallX.Clone. " & ex.Message)
        End Try
    End Function

    Public Class KeyValuePair
        Private cKey As String
        Private cValue As Object

        Public Sub New(ByVal Key As String, ByVal Value As Object)
            cKey = Key
            cValue = Value
        End Sub
        Public Property Key() As String
            Get
                Return cKey
            End Get
            Set(ByVal value As String)
                cKey = value
            End Set
        End Property

        Public Property Value() As Object
            Get
                Return cValue
            End Get
            Set(ByVal value As Object)
                cValue = value
            End Set
        End Property
    End Class

    Public Class CollXError
        Inherits Exception
        Private cMessage As String

        Public Sub New(ByVal Message As String)
            cMessage = Message
        End Sub

        Public Overrides ReadOnly Property Message() As String
            Get
                Return cMessage
            End Get
        End Property

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Class