<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Me.lblProgress = New System.Windows.Forms.Label
        Me.txtProgress = New System.Windows.Forms.TextBox
        Me.lblFile = New System.Windows.Forms.Label
        Me.lblResults = New System.Windows.Forms.Label
        Me.lblContent = New System.Windows.Forms.Label
        Me.lblSearchPath = New System.Windows.Forms.Label
        Me.cboFileOrPattern = New System.Windows.Forms.ComboBox
        Me.lstResults = New System.Windows.Forms.ListBox
        Me.txtSearchFor = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblSearchFor = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnSearch = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.txtSearchPath = New System.Windows.Forms.TextBox
        Me.btnStop = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(-203, 186)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(85, 13)
        Me.lblProgress.TabIndex = 21
        Me.lblProgress.Text = "Search Progress"
        '
        'txtProgress
        '
        Me.txtProgress.Location = New System.Drawing.Point(127, 228)
        Me.txtProgress.Name = "txtProgress"
        Me.txtProgress.Size = New System.Drawing.Size(581, 20)
        Me.txtProgress.TabIndex = 20
        Me.txtProgress.TabStop = False
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(-203, 11)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(103, 13)
        Me.lblFile.TabIndex = 19
        Me.lblFile.Text = "File Name or Pattern"
        '
        'lblResults
        '
        Me.lblResults.AutoSize = True
        Me.lblResults.Location = New System.Drawing.Point(-202, 69)
        Me.lblResults.Name = "lblResults"
        Me.lblResults.Size = New System.Drawing.Size(65, 13)
        Me.lblResults.TabIndex = 18
        Me.lblResults.Text = "Items Found"
        '
        'lblContent
        '
        Me.lblContent.AutoSize = True
        Me.lblContent.Location = New System.Drawing.Point(-202, 44)
        Me.lblContent.Name = "lblContent"
        Me.lblContent.Size = New System.Drawing.Size(65, 13)
        Me.lblContent.TabIndex = 17
        Me.lblContent.Text = "Fille Content"
        '
        'lblSearchPath
        '
        Me.lblSearchPath.AutoSize = True
        Me.lblSearchPath.Location = New System.Drawing.Point(-203, -23)
        Me.lblSearchPath.Name = "lblSearchPath"
        Me.lblSearchPath.Size = New System.Drawing.Size(66, 13)
        Me.lblSearchPath.TabIndex = 16
        Me.lblSearchPath.Text = "Search Path"
        '
        'cboFileOrPattern
        '
        Me.cboFileOrPattern.FormattingEnabled = True
        Me.cboFileOrPattern.Location = New System.Drawing.Point(127, 52)
        Me.cboFileOrPattern.Name = "cboFileOrPattern"
        Me.cboFileOrPattern.Size = New System.Drawing.Size(581, 21)
        Me.cboFileOrPattern.TabIndex = 1
        '
        'lstResults
        '
        Me.lstResults.FormattingEnabled = True
        Me.lstResults.Location = New System.Drawing.Point(127, 118)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.Size = New System.Drawing.Size(581, 95)
        Me.lstResults.TabIndex = 13
        Me.lstResults.TabStop = False
        '
        'txtSearchFor
        '
        Me.txtSearchFor.Location = New System.Drawing.Point(127, 86)
        Me.txtSearchFor.Name = "txtSearchFor"
        Me.txtSearchFor.Size = New System.Drawing.Size(581, 20)
        Me.txtSearchFor.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 228)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "Search Progress"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 13)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "File Name or Pattern"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 118)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Items Found"
        '
        'lblSearchFor
        '
        Me.lblSearchFor.AutoSize = True
        Me.lblSearchFor.Location = New System.Drawing.Point(13, 81)
        Me.lblSearchFor.Name = "lblSearchFor"
        Me.lblSearchFor.Size = New System.Drawing.Size(59, 13)
        Me.lblSearchFor.TabIndex = 23
        Me.lblSearchFor.Text = "Search For"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 14)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Search Path"
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(305, 254)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(125, 76)
        Me.btnSearch.TabIndex = 3
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(584, 254)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(125, 76)
        Me.btnNew.TabIndex = 5
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'txtSearchPath
        '
        Me.txtSearchPath.Location = New System.Drawing.Point(127, 14)
        Me.txtSearchPath.Name = "txtSearchPath"
        Me.txtSearchPath.Size = New System.Drawing.Size(581, 20)
        Me.txtSearchPath.TabIndex = 0
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(445, 254)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(125, 76)
        Me.btnStop.TabIndex = 4
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(759, 369)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.txtSearchPath)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblSearchFor)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.txtProgress)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.lblResults)
        Me.Controls.Add(Me.lblContent)
        Me.Controls.Add(Me.lblSearchPath)
        Me.Controls.Add(Me.cboFileOrPattern)
        Me.Controls.Add(Me.lstResults)
        Me.Controls.Add(Me.txtSearchFor)
        Me.Name = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents txtProgress As System.Windows.Forms.TextBox
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents lblResults As System.Windows.Forms.Label
    Friend WithEvents lblContent As System.Windows.Forms.Label
    Friend WithEvents lblSearchPath As System.Windows.Forms.Label
    Friend WithEvents cboFileOrPattern As System.Windows.Forms.ComboBox
    Friend WithEvents lstResults As System.Windows.Forms.ListBox
    Friend WithEvents txtSearchFor As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblSearchFor As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents txtSearchPath As System.Windows.Forms.TextBox
    Friend WithEvents btnStop As System.Windows.Forms.Button
End Class
