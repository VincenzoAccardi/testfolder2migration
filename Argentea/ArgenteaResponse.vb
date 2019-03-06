Imports System.Collections.Generic

Public Class ArgenteaResponse

    Dim _ReturnCode As ArgenteaFunctionsReturnCode
    Public Property ReturnCode As ArgenteaFunctionsReturnCode
        Get
            Return _ReturnCode
        End Get
        Set(value As ArgenteaFunctionsReturnCode)
            _ReturnCode = value
        End Set
    End Property

    Dim _FunctionType As InternalArgenteaFunctionTypes
    Public Property FunctionType As InternalArgenteaFunctionTypes
        Get
            Return _FunctionType
        End Get
        Set(value As InternalArgenteaFunctionTypes)
            _FunctionType = value
        End Set
    End Property

    Dim _CharSeparator As String
    Public Property CharSeparator As String
        Get
            Return _CharSeparator
        End Get
        Set(value As String)
            _CharSeparator = value
        End Set
    End Property

    Dim _MessageOut As String
    Public Property MessageOut As String
        Get
            Return _MessageOut
        End Get
        Set(value As String)
            _MessageOut = value
        End Set
    End Property

    Dim _TransactionID As String
    Public Property TransactionID As String
        Get
            Return _TransactionID
        End Get
        Set(value As String)
            _TransactionID = value
        End Set
    End Property

    Private _map As New Dictionary(Of String, Object)
    Public Sub SetProperty(name As String, value As Object)
        If Not ExistProperty(name) Then
            _map.Add(name, value)
        Else
            _map(name) = value
        End If
    End Sub
    Public Function GetProperty(name As String) As Object
        If ExistProperty(name) Then
            Return _map(name)
        Else
            Return ""
        End If
    End Function
    Public Function ExistProperty(name As String) As Boolean
        Return _map.ContainsKey(name)
    End Function
End Class

