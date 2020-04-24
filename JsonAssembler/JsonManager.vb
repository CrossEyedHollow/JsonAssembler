﻿Imports System.Net
Imports RestSharp

Class JsonManager

    Private client As RestClient
    Private serverAcc As String
    Private serverPass As String
    Private authType As AuthenticationType
    Private token As AuthenticationToken

    ''' <summary>
    ''' Call this method to initialize the needed internal objects 
    ''' </summary>
    ''' <param name="url"></param>
    Public Sub New(url As String)
        'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
        client = New RestClient(url)
        authType = AuthenticationType.NoAuth
    End Sub
    Public Sub New(url As String, username As String, password As String, authenticationType As AuthenticationType, authToken As AuthenticationToken)
        client = New RestClient(url)
        serverAcc = username
        serverPass = password
        authType = authenticationType
        token = authToken
    End Sub

    Public Function Post(json As String) As IRestResponse
        Dim request As RestRequest = New RestRequest(Method.POST)

        Select Case authType
            Case AuthenticationType.Bearer
                'If there is no valid token avaible return
                If Not token.IsValid Then Throw New Exception("No valid token avaible for the operation")

                'Add the headers and body
                request.AddHeader("cache-control", "no-cache")
                request.AddHeader("Authorization", token.Value)

                'Console.WriteLine("Sending POST with token: " & token.Value)
                request.AddHeader("content-type", "application/json; charset=utf-8")
                request.AddParameter("application/json", json, ParameterType.RequestBody)

            Case AuthenticationType.NoAuth
                'Add the headers and body
                request.AddHeader("cache-control", "no-cache")
                request.AddHeader("Authorization", "Basic Og==")
                request.AddHeader("content-type", "application/json; charset=utf-8")
                request.AddParameter("application/json", json, ParameterType.RequestBody)

            Case Else
                Throw New NotImplementedException($"{authType.ToString()} not implemented yet")
        End Select

        'Execute
        Dim response = client.Execute(request)
        If Not response.IsSuccessful Then
            Throw New Exception($"POST operation failed, error: {response.Content}")
        End If
        Return response
    End Function
End Class

Public Class AuthenticationToken
    Public Property Value As String = ""
    Public Property IsValid As Boolean = False
    Public Property ExpiresIn As Integer = 0
End Class

Public Enum AuthenticationType
    NoAuth
    Basic
    Bearer
End Enum
