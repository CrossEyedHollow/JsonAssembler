Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ReportTools

Public Class JsonListener

    Dim listener As HttpListener
    Shared Property Prefix As String
    Shared Property Users As List(Of User)

    Public Sub New()
        listener = New HttpListener()
        listener.Prefixes.Add("http://localhost:8080/")
        listener.Prefixes.Add("http://127.0.0.1:8080/")
        listener.Prefixes.Add(Prefix)
        listener.AuthenticationSchemes = AuthenticationSchemes.Basic
    End Sub

    ''' <summary>
    ''' Starts the listener
    ''' </summary>
    Public Sub Start()
        Dim thrdListener = New Thread(AddressOf Listen)

        listener.Start()
        thrdListener.Start()
    End Sub

    Private Sub Listen()
        Dim context As HttpListenerContext

        While True
            Try
                'Listen
                context = listener.GetContext()

                'Proccess the message
                Task.Factory.StartNew(Sub() ProccessMessage(context))
            Catch ex As Exception
                'If something fails log the error
                Output.Report($"Failed to process message. Reason: {ex.Message}")
            End Try
        End While
    End Sub

    Public Sub ProccessMessage(context As HttpListenerContext)
        Try
            'Declare variables
            Dim db As New DBManager()
            Dim answer As String = ""
            Dim responseCode As Integer = 202

            'Check credentials
            If CheckCredentials(context) <> AuthResult.Valid Then Exit Sub

            'Convert message to string
            Dim rawText As String = New StreamReader(context.Request.InputStream, context.Request.ContentEncoding).ReadToEnd()

            'Parse the incoming msg
            Dim json As JObject = JObject.Parse(rawText)
            Dim msgType As String = json("Message_Type")
            Dim code As String = json("Code")

            'Process 
            Select Case msgType
                Case "IRU"
                    Dim eventTime As Date = ParseTime(json("Event_Time"))
                    Dim quantity As Integer = json("Req_Quantity")

                    'Save the json in alternative table
                    If db.InsertIRU(rawText, msgType.ToUpper(), eventTime, quantity) Then
                        Output.ToConsole("New IRU received.")
                    End If
                    answer = StandartResponse(code, msgType, rawText.ToMD5Hash(), Nothing, eventTime)
                Case Else
                    'Save the json in alternative table
                    If db.InsertRawJson("tblincomingjson", rawText, msgType.ToUpper()) Then
                        Output.ToConsole("New json received.")
                    End If
                    answer = StandartResponse(code, msgType, rawText.ToMD5Hash(), Nothing)

            End Select

            'Return a response
            context.Respond(answer, responseCode)
        Catch ex As Exception
            Try
                'If something fails, respond with error message and log the error
                Dim reason = New With {.Error_Code = "SYSTEM_ERROR", .Error_Descr = $"Failed to process incoming message, reason: {ex.Message}"}
                Dim response As String = StandartResponse(Nothing, Nothing, Nothing, reason)
                Output.Report(reason.Error_Descr)

                context.Respond(response, 500) '500 = Internal server error
            Catch exx As Exception
                Output.Report($"Failed to respond. Error: {exx.Message}")
            End Try
        End Try
    End Sub

    Public Function CheckCredentials(context As HttpListenerContext) As AuthResult
        Dim output As AuthResult = AuthResult.Valid
        Dim id As HttpListenerBasicIdentity = context.User.Identity
        Dim hashPass As String = ToMD5Hash(id.Password)

        Dim sender = JsonListener.Users.FirstOrDefault(Function(x) x.Name = id.Name)

        'If the user is not found or the password doesnt match
        If sender Is Nothing OrElse sender.Password <> hashPass Then
            ReportTools.Output.Report($"Bad user or password, user: '{id.Name}', pass: '{id.Password}'.")
            output = AuthResult.Invalid
        End If
        Return output
    End Function

    Public Enum AuthResult
        Valid
        Invalid
    End Enum

End Class
