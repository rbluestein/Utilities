Public Class Form1
    Inherits System.Windows.Forms.Form

    Private cFullPath As String
    Private cSelectionStart As Integer
    Private cSelectionLength As Integer
    Private cExcelExecutingInd As Boolean
    Private cNewFindStart As Integer
    Friend WithEvents btnZapExcel As System.Windows.Forms.Button
    Friend WithEvents ddTable As System.Windows.Forms.ComboBox
    Friend WithEvents chkAlwaysOnTop As System.Windows.Forms.CheckBox
    Private cLastSearchFor As String

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnGo As System.Windows.Forms.Button
    Friend WithEvents txtSql As System.Windows.Forms.TextBox
    Friend WithEvents btnNewProd As System.Windows.Forms.Button
    Friend WithEvents btnFormat As System.Windows.Forms.Button
    Friend WithEvents ddDBName As System.Windows.Forms.ComboBox
    Friend WithEvents ddDBHost As System.Windows.Forms.ComboBox
    Friend WithEvents btnSavedQueries As System.Windows.Forms.Button
    Friend WithEvents txtSearchFor As System.Windows.Forms.TextBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnGo = New System.Windows.Forms.Button
        Me.txtSql = New System.Windows.Forms.TextBox
        Me.btnNewProd = New System.Windows.Forms.Button
        Me.btnSavedQueries = New System.Windows.Forms.Button
        Me.btnFormat = New System.Windows.Forms.Button
        Me.ddDBName = New System.Windows.Forms.ComboBox
        Me.ddDBHost = New System.Windows.Forms.ComboBox
        Me.txtSearchFor = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnZapExcel = New System.Windows.Forms.Button
        Me.ddTable = New System.Windows.Forms.ComboBox
        Me.chkAlwaysOnTop = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'btnGo
        '
        Me.btnGo.Location = New System.Drawing.Point(8, 152)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(72, 23)
        Me.btnGo.TabIndex = 2
        Me.btnGo.Text = "Go"
        '
        'txtSql
        '
        Me.txtSql.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSql.Location = New System.Drawing.Point(8, 32)
        Me.txtSql.Multiline = True
        Me.txtSql.Name = "txtSql"
        Me.txtSql.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSql.Size = New System.Drawing.Size(336, 112)
        Me.txtSql.TabIndex = 1
        '
        'btnNewProd
        '
        Me.btnNewProd.Location = New System.Drawing.Point(266, 152)
        Me.btnNewProd.Name = "btnNewProd"
        Me.btnNewProd.Size = New System.Drawing.Size(72, 23)
        Me.btnNewProd.TabIndex = 5
        Me.btnNewProd.Text = "New"
        '
        'btnSavedQueries
        '
        Me.btnSavedQueries.Location = New System.Drawing.Point(180, 152)
        Me.btnSavedQueries.Name = "btnSavedQueries"
        Me.btnSavedQueries.Size = New System.Drawing.Size(72, 23)
        Me.btnSavedQueries.TabIndex = 4
        Me.btnSavedQueries.Text = "Queries"
        '
        'btnFormat
        '
        Me.btnFormat.Location = New System.Drawing.Point(94, 152)
        Me.btnFormat.Name = "btnFormat"
        Me.btnFormat.Size = New System.Drawing.Size(72, 23)
        Me.btnFormat.TabIndex = 3
        Me.btnFormat.Text = "Format"
        '
        'ddDBName
        '
        Me.ddDBName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ddDBName.Location = New System.Drawing.Point(74, 7)
        Me.ddDBName.Name = "ddDBName"
        Me.ddDBName.Size = New System.Drawing.Size(156, 21)
        Me.ddDBName.TabIndex = 6
        '
        'ddDBHost
        '
        Me.ddDBHost.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ddDBHost.Location = New System.Drawing.Point(8, 7)
        Me.ddDBHost.Name = "ddDBHost"
        Me.ddDBHost.Size = New System.Drawing.Size(64, 21)
        Me.ddDBHost.TabIndex = 7
        '
        'txtSearchFor
        '
        Me.txtSearchFor.Location = New System.Drawing.Point(343, 153)
        Me.txtSearchFor.Name = "txtSearchFor"
        Me.txtSearchFor.Size = New System.Drawing.Size(72, 20)
        Me.txtSearchFor.TabIndex = 11
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(420, 152)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(39, 23)
        Me.btnFind.TabIndex = 12
        Me.btnFind.Text = "Find"
        '
        'btnZapExcel
        '
        Me.btnZapExcel.Location = New System.Drawing.Point(396, 8)
        Me.btnZapExcel.Name = "btnZapExcel"
        Me.btnZapExcel.Size = New System.Drawing.Size(63, 20)
        Me.btnZapExcel.TabIndex = 13
        Me.btnZapExcel.Text = "Zap Exl"
        '
        'ddTable
        '
        Me.ddTable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ddTable.Location = New System.Drawing.Point(234, 7)
        Me.ddTable.Name = "ddTable"
        Me.ddTable.Size = New System.Drawing.Size(115, 21)
        Me.ddTable.TabIndex = 14
        '
        'chkAlwaysOnTop
        '
        Me.chkAlwaysOnTop.AutoSize = True
        Me.chkAlwaysOnTop.Location = New System.Drawing.Point(353, 9)
        Me.chkAlwaysOnTop.Name = "chkAlwaysOnTop"
        Me.chkAlwaysOnTop.Size = New System.Drawing.Size(45, 17)
        Me.chkAlwaysOnTop.TabIndex = 15
        Me.chkAlwaysOnTop.Text = "Top"
        Me.chkAlwaysOnTop.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(464, 206)
        Me.Controls.Add(Me.chkAlwaysOnTop)
        Me.Controls.Add(Me.ddTable)
        Me.Controls.Add(Me.btnZapExcel)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.txtSearchFor)
        Me.Controls.Add(Me.ddDBHost)
        Me.Controls.Add(Me.ddDBName)
        Me.Controls.Add(Me.btnFormat)
        Me.Controls.Add(Me.btnSavedQueries)
        Me.Controls.Add(Me.btnNewProd)
        Me.Controls.Add(Me.txtSql)
        Me.Controls.Add(Me.btnGo)
        Me.MinimumSize = New System.Drawing.Size(360, 216)
        Me.Name = "Form1"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.Text = "New Excel Big 6.1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dt As DataTable
        Dim i As Integer
        Dim ReadStream As System.IO.StreamReader
        Dim Box() As String
        Dim IniData As String = Nothing
        Dim DBName As String = Nothing
        Dim DBHost As String = Nothing


        'PassedInArgs = System.Environment.GetCommandLineArgs
        'MessageBox.Show(PassedInArgs.GetUpperBound(0))

        Try

            cFullPath = System.IO.Directory.GetCurrentDirectory() & "\NewExcelBig.ini"

            'Dim excel As New Excel
            'Dim dtExcel As DataTable = excel.ExcelToDT("C:\Users\rbluestein\Desktop")


            ' ___ Read the ini file
            If System.IO.File.Exists(cFullPath) Then
                ReadStream = New System.IO.StreamReader(cFullPath)
                IniData &= ReadStream.ReadToEnd
                ReadStream.Close()
                Box = Split(IniData, "|")
                DBHost = Box(0)
                DBName = Box(1)
            End If

            ' ___ Populate the server dropdown. Select item if available. Set the DBHost
            ddDBHost.Items.Add("hbg-tst")
            ddDBHost.Items.Add("hbg-sql")
            ddDBHost.Items.Add("wadev")
            ddDBHost.Items.Add("training")
            If IniData = Nothing Then
                Common.DBHost = "192.168.1.10"
            Else
                ddDBHost.Text = DBHost
                Common.DBHost = DBHost
            End If


            ' ___ Populate the database dropdown. Select item if available.
            dt = Common.GetDT("SELECT * FROM master..sysdatabases ORDER BY NAME")
            For i = 0 To dt.Rows.Count - 1
                ddDBName.Items.Add(dt.Rows(i)(0))
            Next
            If IniData <> Nothing Then
                ddDBName.Text = DBName
                Common.DBName = DBName
            End If

            PopulateDropdown("ddTable")

        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #330: Form_Load " & ex.Message)
        End Try
    End Sub

    Private Sub SetDBHost(ByVal ServerName As String)
        Select Case ServerName
            Case "wadev"
                'IPAddress = "10.50.1600"
                Common.DBHost = "wadev"
            Case "hbg-tst"
                Common.DBHost = "192.168.1.15"
            Case "hbg-sql"
                Common.DBHost = "192.168.1.10"
            Case "training"
                Common.DBHost = "192.168.1.14"
        End Select
    End Sub

    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
        Try
            If txtSql.Text.Length > 0 Then
                txtSql.Text = Trim(txtSql.Text)
                GoExcel(1)
            End If
        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #331: btnGo_Click " & ex.Message)
        End Try
    End Sub

    'Private Sub txtSql_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSql.KeyUp
    '    If e.KeyValue = 13 And txtSql.Text.Length > 0 Then
    '        txtSql.Text = Trim(txtSql.Text)
    '        Go()
    '    End If
    'End Sub

    Private Sub GoExcel(ByVal Idx As Integer)
        Dim Results As Results
        Dim QueryPack As QueryPack
        Dim Sql As String = Nothing
        Dim Excel As Excel

        Try

            Me.Cursor = Cursors.WaitCursor
            Excel = New Excel
            Common.DBHost = ddDBHost.SelectedItem

            Select Case Idx
                Case 1
                    cSelectionLength = txtSql.SelectionLength
                    cSelectionStart = txtSql.SelectionStart
                    If cSelectionLength > 0 Then
                        Sql = txtSql.Text.Substring(cSelectionStart, cSelectionLength)
                    Else
                        Sql = txtSql.Text
                    End If

                Case 2
                    'Sql = "select table_name from information_schema.tables where table_type='BASE TABLE' order by table_name"
                    Sql = "select table_name, table_type from information_schema.tables order by table_type"
                Case 3
                    'Sql = "select column_name, CHARACTER_MAXIMUM_LENGTH, ordinal_position,column_default,data_type, Is_nullable from information_schema.columns where table_name='" & txtTableName.Text.Trim & "' order by column_name"
                    Sql = "select column_name, CHARACTER_MAXIMUM_LENGTH, ordinal_position,column_default,data_type, Is_nullable from information_schema.columns where table_name='" & ddTable.SelectedItem & "' order by column_name"
            End Select


            Sql = "USE [" & ddDBName.SelectedItem & "] " & Sql

            cExcelExecutingInd = True

            QueryPack = Common.GetDTWithQueryPack(Sql)
            If Not QueryPack.Success Then
                MessageBox.Show("Database error:" & vbCrLf & QueryPack.TechErrMsg)
            End If

            If QueryPack.Success Then
                Results = Excel.ExportToExcel(QueryPack.dt, Nothing)
                If Not Results.Success Then
                    If Results.Msg = "Exception from HRESULT: 0x800A03EC" Then
                        Dim FrmMessage As New Message
                        FrmMessage.Show()
                    Else
                        Select Case txtSql.Text.ToLower.Substring(0, 6)
                            Case "select"
                                MessageBox.Show("Excel error:" & vbCrLf & Results.Msg)
                            Case "insert"
                                If Results.Msg = "Exception from HRESULT: 0x800A03EC." Then
                                    MessageBox.Show("INSERT query completed. Check for accuracy.")
                                End If
                            Case "update"
                                If Results.Msg = "Exception from HRESULT: 0x800A03EC." Then
                                    MessageBox.Show("UPDATE query completed. Check for accuracy.")
                                End If
                        End Select
                    End If
                End If
            End If


            Dim WriteStream As New System.IO.StreamWriter(cFullPath)
            WriteStream.Write(ddDBHost.SelectedItem & "|" & ddDBName.SelectedItem)
            WriteStream.Close()

            'txtSql.SelectionLength = SelectionLength
            'txtSql.SelectionStart = SelectionStart

            Me.Cursor = Cursors.Default

            'Me.Refresh()

        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #332: Go " & ex.Message)
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewProd.Click
        Try
            'Shell("NewExcelBig.exe prod", AppWinStyle.NormalFocus)
            Shell("NewExcelBig.exe", AppWinStyle.NormalFocus)
        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #333: btnNew_Click " & ex.Message)
        End Try
    End Sub

    Private Sub btnSavedQueries_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSavedQueries.Click
        Try
            'Shell("NewExcelBig.exe test", AppWinStyle.NormalFocus)
            Process.Start("notepad.exe", System.IO.Directory.GetCurrentDirectory() & "\SavedQueries.txt")
        Catch ex As Exception
            Report.Report(Report.ReportTypeEnum.Error, "Error #334: btnNewTest_Click " & ex.Message)
        End Try
    End Sub

    Private Sub FrmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        txtSql.Height = Me.Height - 104
        txtSql.Width = Me.Width - 24

        ' original size: 408 by 240

        btnGo.Top = Me.Height - 65
        btnFormat.Top = Me.Height - 65
        btnSavedQueries.Top = Me.Height - 65
        btnNewProd.Top = Me.Height - 65


        btnFind.Top = Me.Height - 65
        txtSearchFor.Top = Me.Height - 65


        'SSTab1.Height = Me.Height - 152
        'treProductionProgs.Height = SSTab1.Height - 32
        'treTestProgs.Height = SSTab1.Height - 32
        'Panel1.Top = SSTab1.Top + SSTab1.Height
    End Sub

    Private Sub btnFormat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormat.Click
        'txtSql.Text = Replace(txtSql.Text, vbCrLf, " ", 1, -1, CompareMethod.Text)

        txtSql.Text = Replace(txtSql.Text, vbCrLf, "")
        txtSql.Text = Replace(txtSql.Text, """", "" & vbCrLf, 1, -1, CompareMethod.Text)
        ' txtSql.Text = Replace(txtSql.Text, "SELECT", vbCrLf & "SELECT" & vbCrLf, 1, -1, CompareMethod.Text)
        txtSql.Text = Replace(txtSql.Text, " SELECT", vbCrLf & "SELECT", 1, -1, CompareMethod.Text)
        txtSql.Text = Replace(txtSql.Text, "FROM", vbCrLf & "FROM", 1, -1, CompareMethod.Text)
        txtSql.Text = Replace(txtSql.Text, "INNER", vbCrLf & "INNER", 1, -1, CompareMethod.Text)
        txtSql.Text = Replace(txtSql.Text, "WHERE", vbCrLf & "WHERE", 1, -1, CompareMethod.Text)
        txtSql.Text = Replace(txtSql.Text, "ORDER BY", vbCrLf & "ORDER BY", 1, -1, CompareMethod.Text)
        txtSql.Text = Replace(txtSql.Text, "AND ", "AND" & vbCrLf, 1, -1, CompareMethod.Text)
    End Sub

    'Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
    '        txtSql.SelectionLength = cSelectionLength
    '        txtSql.SelectionStart = cSelectionStart
    'End Sub

    'Private Sub Form1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MouseDown
    '    txtSql.SelectionLength = cSelectionLength
    '    txtSql.SelectionStart = cSelectionStart
    'End Sub

    'Private Sub txtSql_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSql.GotFocus
    '    txtSql.SelectionLength = cSelectionLength
    '    txtSql.SelectionStart = cSelectionStart
    'End Sub

    'Private Sub txtSql_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtSql.MouseUp
    '    txtSql.SelectionLength = cSelectionLength
    '    txtSql.SelectionStart = cSelectionStart
    'End Sub

    'Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
    '    txtSql.SelectionLength = cSelectionLength
    '    txtSql.SelectionStart = cSelectionStart
    'End Sub

    Private Sub ExcelReturn()
        If cExcelExecutingInd Then

            'txtSql.Text = "abcdefghijklmnopqrstuvwxyz"
            'txtSql.SelectionLength = 4
            'txtSql.SelectionStart = 3

            cExcelExecutingInd = False
            txtSql.SelectionLength = cSelectionLength
            txtSql.SelectionStart = cSelectionStart
            Me.Focus()
            Me.Refresh()
        End If
    End Sub

    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        ExcelReturn()
    End Sub

    Private Sub Form1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
        ExcelReturn()
    End Sub

    Protected Overrides Sub OnGotFocus(ByVal e As System.EventArgs)
        ExcelReturn()
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim FindStart As Integer
        Dim FindLength As Integer
        Dim SearchFor As String

        SearchFor = txtSearchFor.Text

        If (cLastSearchFor = Nothing) Or (SearchFor <> cLastSearchFor) Then
            cLastSearchFor = SearchFor
            cNewFindStart = 1
        End If

        FindStart = InStr(cNewFindStart, txtSql.Text, SearchFor, vbTextCompare)

        If FindStart > 0 Then
            FindLength = Len(SearchFor)
            With txtSql
                .Focus()
                .SelectionStart = FindStart - 1
                .SelectionLength = FindLength
            End With
            cNewFindStart = FindStart + FindLength
            txtSql.ScrollToCaret()
        Else
            MessageBox.Show("Not found.")
        End If


    End Sub

    Private Sub PopulateDropdown(ByVal Name As String)
        Dim i As Integer
        Dim SelectedDB As String = Nothing
        Dim dt As DataTable

        Select Case Name
            Case "ddDBHost"
                PopulateDropdown("ddDBName")

            Case "ddDBName"
                If ddDBName.SelectedItem <> Nothing Then
                    SelectedDB = ddDBName.SelectedItem()
                End If
                ' ___ Does this DB exist on this DBHost?
                dt = Common.GetDT("SELECT COUNT (*) FROM master..sysdatabases WHERE NAME='" & Common.DBName & "'")

                ddDBName.Items.Clear()
                If dt.Rows(0)(0) > 0 Then

                    dt = Common.GetDT("SELECT * FROM master..sysdatabases ORDER BY NAME")
                    For i = 0 To dt.Rows.Count - 1
                        ddDBName.Items.Add(dt.Rows(i)(0))
                        If SelectedDB = dt.Rows(i)(0) Then
                            ddDBName.SelectedIndex = i
                        End If
                    Next
                    PopulateDropdown("ddTable")
                End If

            Case "ddTable"

                ddTable.Items.Clear()

                If Common.DBHost <> Nothing AndAlso Common.DBName <> Nothing Then
                    dt = Common.GetDT("select table_name, table_type from information_schema.tables where table_type = 'base table' order by table_name")
                    For i = 0 To dt.Rows.Count - 1
                        ddTable.Items.Add(dt.Rows(i)(0))
                        If Common.TblName = dt.Rows(i)(0) Then
                            ddTable.SelectedIndex = i
                        End If
                    Next
                End If
        End Select
    End Sub

    Private Sub ddDHost_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddDBHost.SelectedIndexChanged
        Common.DBHost = ddDBHost.SelectedItem
        PopulateDropdown("ddDBHost")
    End Sub

    Private Sub ddDBName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddDBName.SelectedIndexChanged
        Common.DBName = ddDBName.SelectedItem
        PopulateDropdown("ddTable")
    End Sub
    Private Sub ddTable_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddTable.SelectedIndexChanged
        Common.TblName = ddTable.SelectedItem
        GoExcel(3)
    End Sub


    Private Sub btnZapExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZapExcel.Click
        Dim sKillExcel As String
        sKillExcel = "TASKKILL /F /IM Excel.exe"
        Shell(sKillExcel, vbHide)
    End Sub

    Private Sub chkAlwaysOnTop_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAlwaysOnTop.CheckedChanged
        If chkAlwaysOnTop.Checked Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub
End Class
