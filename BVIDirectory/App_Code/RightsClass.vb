Imports System.Data

Public Class RightsClass
    Private cCommon As Common
    Private cEnviro As Enviro
    Private cRightsColl As New Collection

#Region " Constants "
    Public Const DirectoryView As String = "DRV"
#End Region

    Public Enum AccessLevelEnum
        AllAccess = 1
        SupervisorAccess = 2
        EnrollerAccess = 3
    End Enum

    Public ReadOnly Property RightsColl()
        Get
            Return cRightsColl
        End Get
    End Property

    Public Sub New(ByRef Enviro As Enviro, ByVal CurPage As Page)
        Dim dt As DataTable
        Dim i As Integer

        Try

            cEnviro = Enviro
            cCommon = New Common

            dt = cCommon.GetDT("SELECT Role, LocationID FROM Users WHERE UserID ='" & Enviro.LoggedInUserID & "'", cEnviro.DBHost, "UserManagement")
            If dt.Rows.Count = 0 Then
                CurPage.Response.Redirect("InsufficientRights.aspx")
            End If

            Select Case dt.Rows(0)("Role")
                Case "ENROLLER", "IT", "SUPERVISOR", "ADMIN", "ADMIN LIC"
                    cRightsColl.Add(DirectoryView)
                Case "CLIENT"
                    ' none
            End Select

        Catch ex As Exception
            Throw New Exception("Error #1770: RightsClass New. " & ex.Message, ex)
        End Try
    End Sub

    Public Sub GetSecurityFlds(ByRef AccessLevel As AccessLevelEnum, ByRef LoginRole As RoleCatgyEnum)
        Try

            Select Case LoginRole
                Case RoleCatgyEnum.Other
                    AccessLevel = AccessLevelEnum.AllAccess
                Case RoleCatgyEnum.Supervisor
                    AccessLevel = AccessLevelEnum.SupervisorAccess
                Case RoleCatgyEnum.Enroller
                    AccessLevel = AccessLevelEnum.EnrollerAccess
            End Select

        Catch ex As Exception
            Throw New Exception("Error #1771: RightsClass GetSecurityFlds. " & ex.Message, ex)
        End Try
    End Sub


    Public Function HasSufficientRights(ByRef RightsRqd As String(), ByVal RedirectOnError As Boolean, ByRef CurPage As System.Web.UI.Page) As Boolean
        Dim i, j As Integer
        Dim Passed As Boolean

        Try

            For i = 0 To RightsRqd.GetUpperBound(0)
                For j = 1 To cRightsColl.Count
                    If cRightsColl(j) = RightsRqd(i) Then
                        Passed = True
                        Exit For
                    End If
                Next
                If Passed Then
                    Exit For
                End If
            Next

            If Passed Then
                Return True
            Else
                If RedirectOnError Then
                    CurPage.Response.Redirect("InsufficientRights.aspx")
                Else
                    Return False
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Error #1772: RightsClass HasSufficientRights. " & ex.Message, ex)
        End Try
    End Function

    Public Function HasThisRight(ByVal RightCd As String) As Boolean
        Dim i As Integer
        For i = 1 To cRightsColl.Count
            If cRightsColl(i) = RightCd Then
                Return True
            End If
        Next
        Return False
    End Function

End Class

