Imports System.Net
Imports System.Text
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
        Dim byteBody As Byte() = Encoding.UTF8.GetBytes(json)
        Dim hash As String = json.ToMD5Hash()

        Select Case authType
            Case AuthenticationType.Basic
                Dim strBaseCredentials As String = Convert.ToBase64String(Encoding.ASCII.GetBytes(String.Format("{0}:{1}", serverAcc, serverPass)))
                request.AddHeader("Authorization", $"Basic {strBaseCredentials}")

            Case AuthenticationType.Bearer
                'If there is no valid token avaible return
                If Not token.IsValid Then Throw New Exception("No valid token avaible for the operation")

                'Add the headers and body
                request.AddHeader("Authorization", token.Value)

            Case AuthenticationType.NoAuth
                'Add the headers and body
                request.AddHeader("Authorization", "Basic Og==")
            Case Else
                Throw New NotImplementedException($"{authType.ToString()} not implemented yet")
        End Select

        request.AddHeader("Content-Length", byteBody.Length)
        request.AddHeader("X-OriginalHash", hash)
        request.AddHeader("cache-control", "no-cache")
        request.AddHeader("content-type", "application/json; charset=utf-8")
        request.AddParameter("application/json", json, ParameterType.RequestBody)

        'Execute
        Dim response = client.Execute(request)

        Return response
    End Function
End Class

Public Class AuthenticationToken
    Public Property Value As String = ""
    Public Property IsValid As Boolean = False
    Public Property ExpiresIn As Integer = 0
End Class

Public Enum AuthenticationType
    NoAuth = 0
    Basic = 1
    Bearer = 2
End Enum
