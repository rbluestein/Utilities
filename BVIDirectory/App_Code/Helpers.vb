Public Class Results
    Private cSuccess As Boolean
    Private cMsg As String
    Private cResponseAction As ResponseAction
    Private cObtainConfirm As Boolean
    Private cValue As Object
    Private cValue2 As Object

    Public Property Success() As Boolean
        Get
            Return cSuccess
        End Get
        Set(ByVal Value As Boolean)
            cSuccess = Value
        End Set
    End Property
    Public Property Msg() As String
        Get
            Return cMsg
        End Get
        Set(ByVal Value As String)
            cMsg = Value
        End Set
    End Property
    Public Property ResponseAction() As ResponseAction
        Get
            Return cResponseAction
        End Get
        Set(ByVal Value As ResponseAction)
            cResponseAction = Value
        End Set
    End Property
    Public Property ObtainConfirm() As Boolean
        Get
            Return cObtainConfirm
        End Get
        Set(ByVal Value As Boolean)
            cObtainConfirm = Value
        End Set
    End Property
    Public Property Value() As Object
        Get
            Return cValue
        End Get
        Set(ByVal Value As Object)
            cValue = Value
        End Set
    End Property
    Public Property Value2() As Object
        Get
            Return cValue2
        End Get
        Set(ByVal Value As Object)
            cValue2 = Value
        End Set
    End Property
End Class

Public Class Style
    Public Enum StyleType
        NormalEditable = 1
        NoneditableGrayed = 3
        NoneditableWhite = 2
        NotVisible = 4
    End Enum

    Public Shared Sub AddCustomStyle(ByVal tb As TextBox, ByRef Elements As Collection, ByVal Visible As Boolean, ByVal [ReadOnly] As Boolean)
        Dim StyleStr As String
        Dim i As Integer
        For i = 1 To Elements.Count
            StyleStr &= Elements(i).Key & ":" & Elements(i).Value & ";"
        Next
        tb.Attributes.Add("style", StyleStr)
        tb.Visible = Visible
        tb.ReadOnly = [ReadOnly]
    End Sub

    Public Shared Sub AddStyle(ByVal tb As TextBox, ByVal StyleType As StyleType, ByVal Width As Integer, Optional ByVal IsTextArea As Boolean = False)
        If Not IsTextArea Then
            Select Case StyleType
                Case StyleType.NormalEditable
                    tb.Attributes.Add("style", "width:" & Width & ";")
                    tb.Visible = True
                    tb.ReadOnly = False

                Case StyleType.NoneditableGrayed
                    tb.Attributes.Add("style", "width:" & Width & ";border-width:1px;background: #eeeedd;readOnly: true;")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NoneditableWhite
                    tb.Attributes.Add("style", "width:" & Width & ";border-width:1px;border-style:solid;border-color:cccccc;background: #ffffff")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NotVisible
                    tb.Visible = False
            End Select
        Else                                                                ' is a textarea
            Select Case StyleType
                Case StyleType.NormalEditable
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;")
                    tb.Visible = True
                    tb.ReadOnly = False

                Case StyleType.NoneditableGrayed
                    ' tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:hidden;readOnly: true")
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:auto;readOnly: true")
                    tb.Visible = True
                    tb.ReadOnly = True


                Case StyleType.NoneditableWhite
                    'tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;overflow:hidden;readOnly: true")
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;overflow:auto;readOnly: true")
                    tb.Visible = True
                    tb.ReadOnly = True

                Case StyleType.NotVisible
                    tb.Visible = False
            End Select
        End If

    End Sub
    Public Shared Sub AddStyle(ByVal tb As TextBox, ByVal IsVisible As Boolean, ByVal IsReadOnly As Boolean, ByVal Width As Integer, Optional ByVal IsTextArea As Boolean = False)
        tb.Visible = IsVisible
        tb.ReadOnly = IsReadOnly
        If Not IsTextArea Then
            If IsVisible Then
                If IsReadOnly Then
                    tb.Attributes.Add("style", "width:" & Width & ";border-width:1px;background: #eeeedd;readOnly: true;")
                    tb.ReadOnly = True
                Else
                    'tb.Attributes.Add("style", "width:" & Width & ";background: #ffffff;")
                    tb.Attributes.Add("style", "width:" & Width & ";")
                End If
            Else
                ' tb.Attributes.Add("style", "VISIBILITY: hidden")
                tb.Attributes.Add("style", "display: none")
            End If
        Else
            If IsVisible Then
                If IsReadOnly Then
                    'tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:hidden;readOnly: true")
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;overflow:auto;readOnly: true")
                    tb.ReadOnly = True
                Else
                    tb.Attributes.Add("style", "width:" & Width & ";FONT: 10pt Arial, Helvetica, sans-serif;")
                End If
            End If
        End If
    End Sub
End Class

Public Class ErrorArgs
    Private cHeaderMessage As String
    Private cErrorMessage As String

    Public Property HeaderMessage() As String
        Get
            Return cHeaderMessage
        End Get
        Set(ByVal Value As String)
            cHeaderMessage = Value
        End Set
    End Property
    Public Property ErrorMessage() As String
        Get
            Return cErrorMessage
        End Get
        Set(ByVal Value As String)
            cErrorMessage = Value
        End Set
    End Property
End Class

