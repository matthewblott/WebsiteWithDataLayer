Imports System.Data.SqlClient

Partial Public Class Db(Of T)

  Public Shared Function [Get](sql As String, source As Object) As IList(Of T)
    Dim list As New List(Of T)
    For Each row In Db.GetDataTable(sql, source).Rows
      Dim instance = Activator.CreateInstance(GetType(T))
      Db.BindDataRow(row, instance)
      list.Add(instance)
    Next
    Return list
  End Function

  Public Shared Function GetScalar(sql As String, source As Object) As T
    Using conn = Db.GetConnection()
      Using cmd As New SqlCommand(Db.AddSetNoCount(sql), conn)
        Db.AddParams(cmd, source)
        conn.Open()
        Return CType(cmd.ExecuteScalar(), T)
      End Using
    End Using
  End Function

End Class

Partial Public Class Db(Of K, V)
  Public Shared Function GetDictionary(sql As String, source As Object) As IDictionary(Of K, V)
    Dim dictionary As New Dictionary(Of K, V)
    For Each row In Db.GetDataTable(sql, source).Rows
      dictionary.Add(CType(row(0), K), CType(row(1), V))
    Next
    Return dictionary
  End Function

End Class