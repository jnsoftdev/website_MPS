Imports Microsoft.VisualBasic
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient

Module SqlParameterCollectionExtensions
    <Extension()>
    Function AddWithValue(ByVal target As SqlParameterCollection, ByVal parameterName As String, ByVal value As Object, ByVal nullValue As Object) As SqlParameter
        If (value Is Nothing Or value = Nothing) And nullValue = Nothing Then
            Return target.AddWithValue(parameterName, If(nullValue, DBNull.Value))
        End If
        Return target.AddWithValue(parameterName, value)
    End Function

    Public Function isNull(ByVal value As Object) As Object
        Return If(Convert.IsDBNull(value) = Nothing, value, Nothing)
    End Function
End Module