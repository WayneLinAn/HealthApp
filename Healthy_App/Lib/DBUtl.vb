Imports System.Data.Sql
Imports System.Data
Imports System.Data.SqlClient

Public Class DBUtl
    Dim sqlConn As SqlConnection
    Dim sqlAdapter As SqlDataAdapter
    Dim sqlcommand As SqlCommand
    Public conn_ok As Boolean = False

    Private conn_str As String = "Data Source=" & srv_ip & ";Initial Catalog=healthappdb;Persist Security Info=True;User ID=sa;Pwd=hbmlmgr"

    Public Sub set_db_ip(ByVal ip As String)
        'conn_str = "Data Source=" & srv_ip & ";Initial Catalog=hbmldb;Persist Security Info=True;User ID=sa;Pwd=hbmlmgr"
        conn_str = "Data Source= HBMMILL-DEV03" & ";Initial Catalog=healthappdb;Persist Security Info=True;User ID=sa;Pwd=hbmlmgr"
    End Sub

    Public Function conn_open() As Boolean
        Try
            sqlConn = New SqlConnection(conn_str)
            sqlConn.Open()
            conn_ok = True
        Catch ex As Exception
            conn_ok = False
        End Try

        Return conn_ok

    End Function

    Public Function conn_close() As Boolean

        Dim result As Boolean = False

        Try
            sqlConn.Close()
            conn_ok = False
            result = True
        Catch ex As Exception
            conn_ok = False
        End Try

        Return result

    End Function


    Public Function exe_sql(ByVal sql_string As String, ByRef data_ds As DataSet) As Boolean
        Dim result As Boolean = False

        Try

            If conn_ok Then
                sqlcommand = New SqlCommand(sql_string, sqlConn)
                sqlAdapter = New SqlDataAdapter(sqlcommand)
                Dim count As Integer = sqlAdapter.Fill(data_ds)
                sqlAdapter.Dispose()
                sqlcommand.Dispose()
                result = True
            Else
                add_event_log("DB Connect Fail!")
            End If
        Catch ex As Exception
            'MsgBox("SQL_Error!!!""SQL:" & sql_string & vbNewLine & "SQL Exception:" & ex.ToString)
            add_event_log("SQL:" & sql_string)
            add_event_log("SQL Exception:" & ex.ToString)
        End Try
        Return result

    End Function

End Class

