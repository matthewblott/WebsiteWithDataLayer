Imports System.Data
Imports System.Data.SqlClient
Partial Public Class Db

  Public Shared Function GetConnection() As SqlConnection
    Return New SqlConnection(ConfigurationManager.ConnectionStrings("main").ConnectionString)
  End Function

  Public Shared Function AddSetNoCount(sql As String) As String
    Return String.Format("set nocount on {0}", sql)
  End Function

  ' 10 - return, 32 - space, 41 - close bracket, 44 - comma
  Public Shared Sub AddParams(ByRef cmd As SqlCommand, source As Object)
    Dim sql = cmd.CommandText.ToLower()
    For Each p In source.GetType().GetProperties
      Dim has As Func(Of Integer, Boolean) = Function(val) sql.Contains(String.Format("@{0}{1}", p.Name.ToLower(), Char.ConvertFromUtf32(val)))
      If has(10) Or has(32) Or has(41) Or has(44) Then
        Dim value = p.GetValue(source)
        cmd.Parameters.AddWithValue(String.Format("@{0}", p.Name), If(value Is Nothing, DBNull.Value, value))
      End If
    Next
  End Sub

  Public Shared Sub BindSelf(sql As String, source As Object)
    Dim data = Db.GetDataTable(sql, source)
    For Each row In data.Rows
      Db.BindDataRow(row, source)
    Next
  End Sub

  Public Shared Function GetDataTable(sql As String, Optional isStoredProcedure As Boolean = False) As DataTable
    Return Db.GetDataTable(Db.AddSetNoCount(sql), Nothing, isStoredProcedure)
  End Function

  Public Shared Sub Execute(sql As String, source As Object, Optional isStoredProcedure As Boolean = False)
    Using conn = Db.GetConnection()
      Using cmd As New SqlCommand(Db.AddSetNoCount(sql), conn)
        If isStoredProcedure Then
          cmd.CommandType = CommandType.StoredProcedure
        End If
        Db.AddParams(cmd, source)
        conn.Open()
        cmd.ExecuteNonQuery()
      End Using
    End Using
  End Sub

  Public Shared Function GetScalar(sql As String, Optional isStoredProcedure As Boolean = False) As Object
    Return Db.GetScalar(sql, Nothing, isStoredProcedure)
  End Function

  Public Shared Function GetScalar(sql As String, source As Object, Optional isStoredProcedure As Boolean = False) As Object
    Using conn = Db.GetConnection()
      Using cmd As New SqlCommand(If(isStoredProcedure, sql, Db.AddSetNoCount(sql)), conn)
        If isStoredProcedure Then
          cmd.CommandType = CommandType.StoredProcedure
        End If
        If source IsNot Nothing Then
          Db.AddParams(cmd, source)
        End If
        conn.Open()
        Return cmd.ExecuteScalar()
      End Using
    End Using
  End Function

  Public Shared Function GetDataSet(sql As String, source As Object, Optional isStoredProcedure As Boolean = False) As DataSet
    Using conn = Db.GetConnection()
      Using cmd As New SqlCommand(If(isStoredProcedure, sql, Db.AddSetNoCount(sql)), conn)
        If isStoredProcedure Then
          cmd.CommandType = CommandType.StoredProcedure
        End If
        If source IsNot Nothing Then
          Db.AddParams(cmd, source)
        End If
        Using adapter As New SqlDataAdapter(cmd)
          Dim data As New DataSet
          conn.Open()
          adapter.Fill(data)
          Return data
        End Using
      End Using
    End Using
  End Function

  Public Shared Function GetDataTable(sql As String, source As Object, Optional isStoredProcedure As Boolean = False) As DataTable
    Return Db.GetDataSet(sql, source, isStoredProcedure).Tables(0)
  End Function

  Public Shared Sub BindDataRow(row As DataRow, source As Object)
    Dim info = source.GetType().GetProperties()
    For Each p In info
      For Each c In row.Table.Columns
        If p.Name.ToLower() = c.Caption.ToLower() Then
          Dim value = row(c.Caption)
          If p.PropertyType = GetType(Boolean) Then
            value = (value.ToString() = 1.ToString() OrElse value.ToString() = True.ToString())
          End If
          If Object.ReferenceEquals(p.PropertyType.BaseType, GetType([Enum])) Then
            value = Convert.ToInt32(value)
          End If
          If Not DBNull.Value.Equals(value) Then
            p.SetValue(source, value)
          End If
        End If
      Next
    Next
  End Sub

  Public Shared Function CreateSelectQueryFromPoco(source As Object) As String
    Dim sql = String.Empty
    For Each p In source.GetType().GetProperties()
      sql &= p.Name & ", "
    Next
    Return String.Format("select {0} from {1} ", sql.Substring(0, sql.Length - 2), source.GetType().Name)
  End Function

  Public Shared Function CreateUpdateQueryFromPoco(source As Object) As String
    Dim info = source.GetType().GetProperties().Where(Function(prop) prop.Name.ToLower() <> "id"), sql = String.Empty
    For Each p In info
      sql &= String.Format("{0} = @{0}, ", p.Name)
    Next
    sql = String.Format("update {0} set {1} where Id = @id ", source.GetType().Name, sql.Substring(0, sql.Length - 2))
    Return sql
  End Function

  Public Shared Function CreateInsertQueryFromPoco(source As Object) As String
    Dim info = source.GetType().GetProperties().Where(Function(prop) prop.Name.ToLower() <> "id")
    Dim fields = String.Empty, values = String.Empty
    For Each p In info
      fields &= p.Name & ", "
      values &= "@" & p.Name & ", "
    Next
    fields = fields.Substring(0, fields.Length - 2)
    values = values.Substring(0, values.Length - 2)
    Return String.Format("insert into {0} ({1}) values ({2}) select scope_identity() ", source.GetType().Name, fields, values)
  End Function

  Public Shared Function CreateDeleteQueryFromPoco(source As Object) As String
    Return String.Format("delete from {0} where Id = @id ", source.GetType().Name)
  End Function

End Class