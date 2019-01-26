Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Public Module CRUD
    <Extension()>
    Public Function Create(ByVal CommandSql As SqlCommand)
        Return ExecuteCommand(CommandSql)
    End Function
    <Extension()>
    Public Function Delete(ByVal CommandSql As SqlCommand)
        Return ExecuteCommand(CommandSql)
    End Function
    <Extension()>
    Public Function Update(ByVal CommandSql As SqlCommand)
        Return ExecuteCommand(CommandSql)
    End Function
    <Extension()>
    Public Function Read(ByVal CommandSql As SqlCommand)
        Try
            If Open() Then
                CommandSql.Connection = ConnectionSql
                If CommandSql.ExecuteNonQuery Then
                    Dim dataTable As New DataTable
                    Dim sqlDataAdapter As New SqlDataAdapter(CommandSql)
                    sqlDataAdapter.Fill(dataTable)
                    If IsDBNull(dataTable) Then
                        Return Nothing
                    Else
                        Return dataTable
                    End If
                Else
                    Return New Exception("No se pudo ejecutar el comando Sql")
                End If
            Else
                Return New Exception("No se pudo abrir la coneccion a la base")
            End If
        Catch ex As Exception
            Return ex
        Finally
            Close()
        End Try
    End Function
    Private Function ExecuteCommand(ByVal CommandSql As SqlCommand)
        Try
            If Open() Then
                CommandSql.Connection = ConnectionSql
                If CommandSql.ExecuteNonQuery Then
                    Return New Exception("El comando Sql fue ejecutado")
                Else
                    Return New Exception("No se pudo ejecutar el comando Sql")
                End If
            Else
                Return New Exception("No se pudo abrir la coneccion a la base")
            End If
        Catch ex As Exception
            Return ex
        Finally
            Close()
        End Try
    End Function
End Module
