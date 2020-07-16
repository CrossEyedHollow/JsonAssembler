Imports System.Runtime.CompilerServices
Imports System.Text
Imports Newtonsoft.Json.Linq

Public Module Extenders
    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As String) As Boolean
        Return IsDBNull(array) OrElse array Is Nothing OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As Integer) As Boolean
        Return IsDBNull(array) OrElse array Is Nothing OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As Decimal) As Boolean
        Return IsDBNull(array) OrElse array Is Nothing OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal str As String) As Boolean
        Return String.IsNullOrEmpty(str) OrElse str Is Nothing
    End Function

    '<Extension()>
    'Public Function ToJSON(ByVal str As String) As String
    '    Return str.Replace("""", "\""")
    'End Function

    ''' <summary>
    ''' Returns all rolls of a single column as a string array
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="columnName"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function ColumnToArray(ByVal str As DataTable, columnName As String) As String()
        Return str.Rows.OfType(Of DataRow).Select(Function(dr) dr.Field(Of String)(columnName)).ToArray()
    End Function

    <Extension()>
    Public Function ToJArray(ByRef array As Array) As JArray
        If IsDBNull(array) Then Return Nothing
        If array Is Nothing Then Return Nothing
        If array.Length = 0 Then Return Nothing

        Return JArray.FromObject(array)
    End Function

    <Extension()>
    Public Function ToMD5Hash(ByVal input As String) As String
        Using md5 As Security.Cryptography.MD5 = Security.Cryptography.MD5.Create()
            'Get the bytes
            Dim inputBytes As Byte() = Encoding.ASCII.GetBytes(input)
            'Compute the hash
            Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)
            Dim sb As StringBuilder = New StringBuilder()
            'Convert to string
            For i As Integer = 0 To hashBytes.Length - 1
                sb.Append(hashBytes(i).ToString("x2"))
            Next
            Return sb.ToString()
        End Using
    End Function
End Module
