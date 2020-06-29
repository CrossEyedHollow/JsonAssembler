Imports System.Runtime.CompilerServices
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

    '<Extension()>
    'Public Function ToJArray(ByRef array As String()) As JArray
    '    If IsDBNull(array) Then Return "null"
    '    If array Is Nothing Then Return "null"
    '    If array.Length = 0 Then Return "null"

    '    Return JArray.FromObject(array)
    'End Function

    '<Extension()>
    'Public Function ToJArray(ByRef array As Decimal()) As JArray
    '    If IsDBNull(array) Then Return "null"
    '    If array Is Nothing Then Return "null"
    '    If array.Length = 0 Then Return "null"

    '    Return JArray.FromObject(array)
    'End Function

    <Extension()>
    Public Function ToJArray(ByRef array As Array) As JArray
        If IsDBNull(array) Then Return Nothing
        If array Is Nothing Then Return Nothing
        If array.Length = 0 Then Return Nothing

        Return JArray.FromObject(array)
    End Function
End Module
