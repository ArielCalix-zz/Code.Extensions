Imports System.Data.SqlClient
Module ConnectionManager
    Friend ConnectionSql As New SqlConnection(Configuration.ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
    Friend Function Open()
        Try
            If ConnectionSql.State = ConnectionState.Open Then
                Return True
            Else
                ConnectionSql.Open()
                Return True
            End If
        Catch ex As Exception
            MsgBox("Error al conectarse con la base de datos" + vbCrLf + ex.Message)
            Return False
        End Try
    End Function
    Friend Function Close()
        Try
            If ConnectionSql.State = ConnectionState.Open Then
                ConnectionSql.Close()
                Return True
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox("Error al conectar con la base de datos" + vbCrLf + ex.Message)
            Return False
        End Try
    End Function
End Module