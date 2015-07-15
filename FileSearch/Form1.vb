Imports System.IO

Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lstResults.HorizontalScrollbar = True

        ' ___ If launched from the Send To menu, use the argument as our directory
        If My.Application.CommandLineArgs.Count > 0 Then
            txtSearchPath.Text = My.Application.CommandLineArgs(0)
            txtContent.Focus()
        End If
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        ShowProgress(txtSearchPath.Text)
        lstResults.Items.Clear()
        Dim Dir_Info As New DirectoryInfo(txtSearchPath.Text)
        ProcessSearch(lstResults, cboFile.Text, Dir_Info, txtContent.Text)
        'Beep()
        txtProgress.Text = "Done"
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        'Shell("NewExcelBig.exe prod", AppWinStyle.NormalFocus)
        Shell("FileSearch.exe", AppWinStyle.NormalFocus)
    End Sub

    Private Sub ShowProgress(ByVal Status As String)
        txtProgress.Text = Status
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub ProcessSearch(ByVal lstResults As ListBox, ByVal FileName As String, ByVal Dir_Info As DirectoryInfo, ByVal Content As String)
        Dim FileCompleteText As String
        Dim _File_Info() As FileInfo
        Dim _SubDir() As DirectoryInfo

        ' ___  Get the files in this directory
        ' Dim fs_infos() As FileInfo = dir_info.GetFiles()
        _File_Info = Dir_Info.GetFiles(FileName)

        ' ___ No content specified: display full path. Otherwise, search against content.
        For Each File_Info As FileInfo In _File_Info
            If Content.Length = 0 Then
                lstResults.Items.Add(File_Info.FullName)
            Else
                FileCompleteText = File.ReadAllText(File_Info.FullName)
                If FileCompleteText.IndexOf(Content, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    lstResults.Items.Add(File_Info.FullName)
                End If
            End If
        Next File_Info
        _File_Info = Nothing

        _SubDir = Dir_Info.GetDirectories()
        For Each SubDir As DirectoryInfo In _SubDir
            ShowProgress(SubDir.FullName)
            ProcessSearch(lstResults, FileName, SubDir, Content)
        Next SubDir
    End Sub
End Class
