Public Class Report
    Private cEnviro As Enviro
    Private cCommon As Common

    Public Sub New()
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        cEnviro = SessionObj("Enviro")
        cCommon = New Common
    End Sub

    Public Function Report(ByVal ErrorMessage As String, ByVal WriteToLog As Boolean, ByVal SendEmailPlease As Boolean, ByVal Shutdown As Boolean) As String
        Dim Coll As New Collection
        Dim SendEmailResults As Results
        Dim HeaderMessage As String = String.Empty
        Dim EmailSuccess As Boolean
        Dim Subj As String
        Dim Textbody As String

        Try

            ' ___ Do not send emails or make log entries for errors occurring on development machine.
            If Environment.MachineName = "LT-5ZFYRC1" Then
                SendEmailPlease = False
                WriteToLog = False
            End If

            ' ___ Write to log
            If WriteToLog Then
                If Shutdown Then
                    WriteToLogFile(ErrorMessage & " ** ERROR FORCING APPLICATION SHUTDOWN **")
                Else
                    WriteToLogFile(ErrorMessage)
                End If
            End If

            ' ___ Send email
            If SendEmailPlease Then
                Coll.Add(cEnviro.LogFileFullPath)
                Subj = "BVI Directory error - User: " & IIf(cEnviro.LoggedInUserName = Nothing, cEnviro.LoggedInUserID, cEnviro.LoggedInUserName) & " - Time: " & cCommon.GetServerDateTime()
                SendEmailResults = SendEmail("HelpDesk@benefitvision.com;rbluestein@benefitvision.com", "automail@benefitvision.com", "", Subj, ErrorMessage)
                EmailSuccess = SendEmailResults.Success
            End If

            If Shutdown And WriteToLog And EmailSuccess Then
                HeaderMessage = "An error has occurred requiring BVI Directory to shut down. BVI Directory has emailed a notice to the help desk. You may view the details of the problem in the log."
            ElseIf Shutdown And WriteToLog And (Not EmailSuccess) Then
                HeaderMessage = "An error has occurred requiring BVI Directory to shut down. You may view the details of the problem in the log."
            ElseIf Shutdown And (Not WriteToLog) And EmailSuccess Then
                HeaderMessage = "An error has occurred requiring BVI Directory to shut down. BVI Directory has emailed a notice to the help desk."
            ElseIf Shutdown And (Not WriteToLog) And (Not EmailSuccess) Then
                HeaderMessage = "An error has occurred requiring BVI Directory to shut down. "
            End If

            Return HeaderMessage.Trim

        Catch ex As Exception
            Throw New Exception("Error #1502: Report Report. " & ex.Message, ex)
        End Try
    End Function

    Public Function SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String, Optional ByRef AttachmentColl As Collection = Nothing) As Results
        Dim schema As String
        Dim MyResults As New Results
        Dim i As Integer
        Dim CDOConfig As CDO.Configuration
        Dim iMsg As CDO.Message

        Try

            '''' ___ Email addresses
            '''SendTo = FeedAdminRow("Developer") & "@benefitvision.com, jkleiman@benefitvision.com"
            '''SentFrom = "automail@benefitvision.com"
            '''cc = "rbluestein@benefitvision.com"

            'Dim CDOConfig As New CDO.Configuration
            CDOConfig = New CDO.Configuration

            schema = "http://schemas.microsoft.com/cdo/configuration/"
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2
            ' CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "mail.benefitvision.com"

            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").Value = 1
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername").Value = "automail"
            CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword").Value = "$bambam2004#"

            CDOConfig.Fields.Update()

            'Dim iMsg As New CDO.Message
            iMsg = New CDO.Message
            iMsg.To = SendTo
            iMsg.From = From
            iMsg.CC = cc
            iMsg.Subject = Subject

            If Not AttachmentColl Is Nothing Then
                For i = 1 To AttachmentColl.Count
                    iMsg.AddAttachment(AttachmentColl(i))
                Next
            End If

            iMsg.Configuration = CDOConfig
            iMsg.TextBody = TextBody
            'imsg.HTMLBody = htmlbody

            '  iMsg.Send()

            '' ___ Clean up
            'iMsg.Attachments.DeleteAll()
            'CDOConfig = Nothing
            'iMsg = Nothing

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Msg = "Error #1503: " & ex.Message
            Return MyResults

        Finally
            iMsg.Attachments.DeleteAll()
            CDOConfig = Nothing
            iMsg = Nothing
        End Try
    End Function



    '		System.Web.Mail.MailMessage Mail = new System.Web.Mail.MailMessage();
    '		Mail.Fields["http://schemas.microsoft.com/cdo/configuration/smtpserver"] = sHost;   
    '		Mail.Fields["http://schemas.microsoft.com/cdo/configuration/sendusing"] = 2;   

    '		Mail.Fields["http://schemas.microsoft.com/cdo/configuration/smtpserverport"] = nPort.ToString();    

    '		if( sUserName.Length == 0 )
    '		{
    '			//Ingen auth
    '		}
    '		else
    '		{
    '			Mail.Fields["http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"] = 1;    
    '			Mail.Fields["http://schemas.microsoft.com/cdo/configuration/sendusername"] = sUserName;     
    '			Mail.Fields["http://schemas.microsoft.com/cdo/configuration/sendpassword"] = sPassword; 
    '		}

    '		Mail.To = sToEmail;   
    '		Mail.From = sFromEmail;   
    '		Mail.Subject = sHeader;   
    '		Mail.Body = sMessage;
    '		Mail.BodyFormat = System.Web.Mail.MailFormat.Html;

    '		System.Web.Mail.SmtpMail.SmtpServer = sHost;   
    '		System.Web.Mail.SmtpMail.Send(Mail);
    '}


    'Public Function SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String, Optional ByRef AttachmentColl As Collection = Nothing) As Results
    '    Dim schema As String
    '    Dim MyResults As New Results
    '    Dim i As Integer

    '    Try

    '        Dim CDOConfig As New CDO.Configuration
    '        schema = "http://schemas.microsoft.com/cdo/configuration/"
    '        CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2
    '        ' CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
    '        CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "mail.benefitvision.com"
    '        CDOConfig.Fields.Update()

    '        Dim iMsg As New CDO.Message
    '        iMsg.To = SendTo
    '        iMsg.From = From
    '        iMsg.CC = cc
    '        iMsg.Subject = Subject

    '        If Not AttachmentColl Is Nothing Then
    '            For i = 1 To AttachmentColl.Count
    '                iMsg.AddAttachment(AttachmentColl(i))
    '            Next
    '        End If

    '        iMsg.Configuration = CDOConfig
    '        iMsg.TextBody = TextBody
    '        'imsg.HTMLBody = htmlbody

    '        iMsg.Send()

    '        ' ___ Clean up
    '        iMsg.Attachments.DeleteAll()
    '        CDOConfig = Nothing
    '        iMsg = Nothing

    '        MyResults.Success = True
    '        Return MyResults

    '    Catch ex As Exception
    '        MyResults.Success = False
    '        MyResults.Msg = "Error #1503: " & ex.Message
    '        Return MyResults
    '    End Try
    'End Function

    Private Function ReadLogFile() As String
        Dim StreamReader As System.IO.StreamReader
        Dim FileText As String

        Try
            StreamReader = New System.IO.StreamReader(cEnviro.LogFileFullPath)
            FileText = StreamReader.ReadToEnd

            'Do While StreamReader.Peek() >= 0
            '    'Console.WriteLine(StreamReader.ReadLine())

            '    x = StreamReader.ReadLine()
            'Loop
            'StreamReader.Close()
            Return FileText

        Catch ex As Exception
            Throw New Exception("Error #1504: Report ReadLogFile. " & ex.Message, ex)
        Finally
            Try
                StreamReader.Close()
            Catch
            End Try
        End Try
    End Function

    Public Sub WriteToLogFile(ByVal Message As String)
        Dim i As Integer
        'Dim FileInfo As System.IO.FileInfo
        Dim StreamWriter As System.IO.StreamWriter

        Try

            Message = Replace(Message, "~", "")

            'FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)

            Try
                StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
            Catch
                Dim procList() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                For i = 0 To procList.GetUpperBound(0)
                    If procList(i).ProcessName = "notepad" Then
                        procList(i).Kill()
                    End If
                    'System.Diagnostics.Debug.WriteLine(procList(k).ProcessName)
                Next
                StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
            End Try



            'StreamWriter.Write("[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] " & Message & vbCrLf)
            StreamWriter.Write(GetTimeStamp() & Message & vbCrLf)
            ' StreamWriter.Close()
        Catch ex As Exception
            'Throw New Exception("Error #1505: Report WriteToLogFile. " & ex.Message, ex)
        Finally
            Try
                StreamWriter.Close()
            Catch
            End Try
        End Try
    End Sub

    Private Function GetTimeStamp() As String
        Return "[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & " " & cEnviro.LoggedInUserID & "] "
    End Function
End Class
