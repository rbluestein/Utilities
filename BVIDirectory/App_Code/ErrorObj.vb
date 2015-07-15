Public Class ErrorObj

    ' //////////////////////////////////////////////////
    ' Why the try try you ask.
    ' An exception includes a message parameter. In
    ' order to include the error message with a number and 
    ' method name the next higher exception in the hierarchy 
    ' prefixes this information.
    '
    ' If an error occurs in the top-most exception level
    ' in the hierarchy, the only way of including both
    ' the exception message and the prefix message is
    ' by throwing an additional exception for the 
    ' prefix message.
    '
    ' Microsoft bug alert
    ' Error handling uses a session object, called
    ' ErrorArgs. If the error handling is written such
    ' that Global.Application_Error is used at all,
    ' it causes the loss of the session objects. This
    ' application processes errors in the ErrorObj 
    ' as an alternative.
    ' //////////////////////////////////////////////////

#Region " Declarations "
    ' Private cErrorType As ErrorTypeEnum
    ' Private cErrorMessage As String
#End Region

#Region " Enums "
    Public Enum ErrorTypeEnum
        [Error] = 1
        Cookie = 2
        Timeout = 3
        Connection = 4
        ProductRegistration = 5
    End Enum
#End Region

#Region " Properties "
    'Public ReadOnly Property ErrorType() As ErrorTypeEnum
    '    Get
    '        Return cErrorType
    '    End Get
    'End Property
    'Public ReadOnly Property ErrorMessage() As String
    '    Get
    '        Return cErrorMessage
    '    End Get
    'End Property
#End Region

#Region " Methods "
    Public Sub New(ByRef Exception As Exception)
        Dim CurException As Exception
        Dim Coll As New Collection
        Dim ExceptionMessage As String

        Try

            ' ___ Extract the error from the exception object
            CurException = Exception
            While Not (CurException Is Nothing)
                Coll.Add(CurException.Message)
                CurException = CurException.InnerException
            End While
            ExceptionMessage = Coll(Coll.Count - 1)

            If InStr(ExceptionMessage, "Thread was being aborted.") > 0 Then
                Exit Sub
            End If

            HandleError(ExceptionMessage, True)
        Catch ex As Exception
            'Throw New Exception("Error #2302: ErrorObj New. " & ex.Message, ex)
        End Try
    End Sub

    Public Sub New(ByRef ErrorMessage As String)
        Try
            HandleError(ErrorMessage, False)
        Catch ex As Exception
        End Try
    End Sub


    Public Sub HandleError(ByVal RawMessage As String, ByVal Shutdown As Boolean)
        Dim Enviro As Enviro
        Dim ErrorType As ErrorTypeEnum
        Dim ReportType As ReportTypeEnum
        Dim HeaderMessage As String
        Dim ErrorArgs As ErrorArgs
        Dim ErrorMessage As String
        Dim WriteToLog As Boolean
        Dim SendEmail As Boolean
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Dim Report As New Report

        Try

            ' ___ Get the response object
            Dim Response As System.Web.HttpResponse
            Response = HttpContext.Current.Response

            ' ___ Get Enviro from session
            Enviro = SessionObj("Enviro")

            ' ___ Get the ErrorArgs from session
            ErrorArgs = SessionObj("ErrorArgs")

            ' ___ Get the ErrorType and ErrorMessage
            If InStr(RawMessage, "UnableToConnect", CompareMethod.Text) > 0 Then
                ReportType = ReportTypeEnum.Error
                ErrorMessage = "BVI Directory is unable to establish a connection with " & Enviro.DBHost & ".  You will not be able to view the BVI Directory until the database connection is restored."
                WriteToLog = True
                SendEmail = False
                Shutdown = True
            ElseIf InStr(RawMessage, "timed out", CompareMethod.Text) > 0 Then
                ReportType = ReportTypeEnum.Timeout
                ErrorMessage = RawMessage
                WriteToLog = False
                SendEmail = False
                Shutdown = True
            Else
                ErrorType = ErrorTypeEnum.Error
                ErrorMessage = RawMessage
                WriteToLog = True
                SendEmail = True
                Shutdown = True
            End If

            ' ___ Send the email
            HeaderMessage = Report.Report(ErrorMessage, WriteToLog, SendEmail, Shutdown)
            ErrorArgs.HeaderMessage = HeaderMessage
            ErrorArgs.ErrorMessage = ErrorMessage

            ' ___ Handle possible redirect to error page and application shutdown.
            If Shutdown Then
                ErrorMessage = Replace(ErrorMessage, "#", "[sharp]")
                ErrorMessage = Replace(ErrorMessage, vbCrLf, "~")
                ErrorMessage = Replace(ErrorMessage, Chr(10), "~")
                If ErrorMessage.Length + HeaderMessage.Length > 2000 Then
                    ErrorMessage = ErrorMessage.Substring(0, (2000 - HeaderMessage.Length))
                End If

                Response.Redirect("ErrorPage.aspx?ErrorMessage=" & ErrorMessage & "&HeaderMessage=" & HeaderMessage)
            End If

        Catch ex As Exception
            'Throw New Exception("Error #2303: ErrorObj HandleError. " & ex.Message, ex)
        End Try
    End Sub
#End Region
End Class

