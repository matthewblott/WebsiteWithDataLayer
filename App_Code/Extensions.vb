Imports System.Runtime.CompilerServices
Imports System.Data

Public Module Extensions
  <Extension()>
  Public Function ToSafeString(value As String) As Object
    If value Is Nothing Then
      Return DBNull.Value
    Else
      Return value
    End If
  End Function

  <Extension()>
  Public Function ToSafeInt(value As Integer) As Object
    If value = 0 Then
      Return DBNull.Value
    Else
      Return value
    End If
  End Function

  <Extension()>
  Public Function ToSafeDate(value As Date) As Object
    If value = Date.MinValue Then
      Return DBNull.Value
    Else
      Return value
    End If
  End Function

  <Extension()>
  Public Function ToCurrency(value As Decimal) As String
    Return String.Format("{0:c}", value)
  End Function

End Module