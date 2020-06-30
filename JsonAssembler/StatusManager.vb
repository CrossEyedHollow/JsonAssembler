Imports System.Threading
Imports Newtonsoft.Json.Linq

Public Class StatusManager
    Dim thrdListener As Thread
    Public URL As String

    Private serverAcc As String
    Private serverPass As String
    Private authType As AuthenticationType
    Private token As AuthenticationToken

    Public Sub New(statusURL As String)
        URL = statusURL
    End Sub

    Public Sub New(url As String, username As String, password As String, authenticationType As AuthenticationType, authToken As AuthenticationToken)
        Me.URL = url
        serverAcc = username
        serverPass = password
        authType = authenticationType
        token = authToken
    End Sub

    Public Sub Start()
        Dim thrdListener = New Thread(AddressOf CheckStatus)
        thrdListener.Start()
    End Sub

    Public Sub Abort()
        thrdListener.Abort()
    End Sub

    Public Sub CheckStatus()
        Dim db As DBManager = New DBManager()
        Dim jMan As JsonManager = New JsonManager(URL, serverAcc, serverPass, authType, token)

        While True
            'Check for jsons with fldStatus = 0
            Dim unconfirmed As DataTable = db.CheckJsonStatus()
            'If any are found
            If unconfirmed.Rows.Count > 0 Then
                For Each json As DataRow In unconfirmed.Rows
                    Try
                        'Get the code
                        Dim index As Integer = Convert.ToInt32(json("fldIndex"))
                        Dim code As String = json("fldRecallCode")
                        'Create STA type of Json
                        Dim jsonBody As String = JsonOperationals.STA(code)
                        'Send it
                        Dim response As String = jMan.Post(jsonBody).Content

                        'Update database
                        Dim jResponse As JObject = JObject.Parse(response)
                        Dim errors As Integer = Convert.ToInt32(jResponse("Error"))
                        Dim errorArr As String = jResponse("Errors").ToString()

                        db.UpdateStatus(index, errors, errorArr)
                        ReportTools.Output.ToConsole($"json status updated at index: {index}, errors: {errors}")
                    Catch ex As Exception
                        ReportTools.Output.Report($"STA message fail: {ex.Message}")
                    End Try

                Next
            End If

            Thread.Sleep(5000)
        End While
    End Sub

End Class
