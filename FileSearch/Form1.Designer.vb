<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtSearchPath = New System.Windows.Forms.TextBox
        Me.txtContent = New System.Windows.Forms.TextBox
        Me.lstResults = New System.Windows.Forms.ListBox
        Me.cboFile = New System.Windows.Forms.ComboBox
        Me.btnSearch = New System.Windows.Forms.Button
        Me.lblSearchPath = New System.Windows.Forms.Label
        Me.lblContent = New System.Windows.Forms.Label
        Me.lblResults = New System.Windows.Forms.Label
        Me.lblFile = New System.Windows.Forms.Label
        Me.txtProgress = New System.Windows.Forms.TextBox
        Me.lblProgress = New System.Windows.Forms.Label
        Me.btnNew = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'txtSearchPath
        '
        Me.txtSearchPath.Location = New System.Drawing.Point(126, 16)
        Me.txtSearchPath.Name = "txtSearchPath"
        Me.txtSearchPath.Size = New System.Drawing.Size(581, 20)
        Me.txtSearchPath.TabIndex = 0
        '
        'txtContent
        '
        Me.txtContent.Location = New System.Drawing.Point(126, 83)
        Me.txtContent.Name = "txtContent"
        Me.txtContent.Size = New System.Drawing.Size(581, 20)
        Me.txtContent.TabIndex = 2
        '
        'lstResults
        '
        Me.lstResults.FormattingEnabled = True
        Me.lstResults.Location = New System.Drawing.Point(126, 115)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.Size = New System.Drawing.Size(581, 95)
        Me.lstResults.TabIndex = 2
        Me.lstResults.TabStop = False
        '
        'cboFile
        '
        Me.cboFile.FormattingEnabled = True
        Me.cboFile.Location = New System.Drawing.Point(126, 49)
        Me.cboFile.Name = "cboFile"
        Me.cboFile.Size = New System.Drawing.Size(581, 21)
        Me.cboFile.TabIndex = 1
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(582, 260)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(125, 76)
        Me.btnSearch.TabIndex = 3
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'lblSearchPath
        '
        Me.lblSearchPath.AutoSize = True
        Me.lblSearchPath.Location = New System.Drawing.Point(17, 23)
        Me.lblSearchPath.Name = "lblSearchPath"
        Me.lblSearchPath.Size = New System.Drawing.Size(66, 13)
        Me.lblSearchPath.TabIndex = 5
        Me.lblSearchPath.Text = "Search Path"
        '
        'lblContent
        '
        Me.lblContent.AutoSize = True
        Me.lblContent.Location = New System.Drawing.Point(18, 90)
        Me.lblContent.Name = "lblContent"
        Me.lblContent.Size = New System.Drawing.Size(65, 13)
        Me.lblContent.TabIndex = 6
        Me.lblContent.Text = "Fille Content"
        '
        'lblResults
        '
        Me.lblResults.AutoSize = True
        Me.lblResults.Location = New System.Drawing.Point(18, 115)
        Me.lblResults.Name = "lblResults"
        Me.lblResults.Size = New System.Drawing.Size(65, 13)
        Me.lblResults.TabIndex = 7
        Me.lblResults.Text = "Items Found"
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(17, 57)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(103, 13)
        Me.lblFile.TabIndex = 8
        Me.lblFile.Text = "File Name or Pattern"
        '
        'txtProgress
        '
        Me.txtProgress.Location = New System.Drawing.Point(126, 225)
        Me.txtProgress.Name = "txtProgress"
        Me.txtProgress.Size = New System.Drawing.Size(581, 20)
        Me.txtProgress.TabIndex = 9
        Me.txtProgress.TabStop = False
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(17, 232)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(85, 13)
        Me.lblProgress.TabIndex = 10
        Me.lblProgress.Text = "Search Progress"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(438, 260)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(125, 76)
        Me.btnNew.TabIndex = 11
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(724, 355)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.txtProgress)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.lblResults)
        Me.Controls.Add(Me.lblContent)
        Me.Controls.Add(Me.lblSearchPath)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.cboFile)
        Me.Controls.Add(Me.lstResults)
        Me.Controls.Add(Me.txtContent)
        Me.Controls.Add(Me.txtSearchPath)
        Me.Name = "Form1"
        Me.Text = "File Search 1.00"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtSearchPath As System.Windows.Forms.TextBox
    Friend WithEvents txtContent As System.Windows.Forms.TextBox
    Friend WithEvents lstResults As System.Windows.Forms.ListBox
    Friend WithEvents cboFile As System.Windows.Forms.ComboBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents lblSearchPath As System.Windows.Forms.Label
    Friend WithEvents lblContent As System.Windows.Forms.Label
    Friend WithEvents lblResults As System.Windows.Forms.Label
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents txtProgress As System.Windows.Forms.TextBox
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents btnNew As System.Windows.Forms.Button

End Class
