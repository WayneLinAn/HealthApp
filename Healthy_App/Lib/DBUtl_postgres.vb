Imports System.Data.Odbc
Imports System.Data.Odbc.OdbcConnection
Imports System.Threading

Public Class DBUtl_postgres


    'RT
    '  Private srv_ip As String = "127.0.0.1"
    ' Private srv_ip As String = "192.168.17.247"
    ' Private srv_ip As String = "192.168.17.248"
    Private conn_str As String = "Driver={PostgreSQL Unicode};Server=" & srv_ip & ";Port=5432;Database=hbmldb;Uid=hbmlmgr;Pwd=hbmlmgr;"

    'Private conn_str As String = "Driver={PostgreSQL ANSI};Server=" & srv_ip & ";Port=5432;Database=hbmldb;Uid=hbmlmgr;Pwd=hbmlmgr;"
    'Private conn_str As String = "Driver={PostgreSQL ODBC Driver(ANSI)};Server=" & srv_ip & ";Port=5432;Database=hbmldb;Uid=hbmlmgr;Pwd=hbmlmgr;"

    Private odbc_conn As OdbcConnection
    Public conn_ok As Boolean = False
    Dim cnt As Integer = 2

    Public Sub set_db_ip(ByVal ip As String)
        conn_str = "Driver={PostgreSQL Unicode};Server=" & ip & ";Port=5432;Database=hbmldb;Uid=hbmlmgr;Pwd=hbmlmgr;"
    End Sub

    Public Function conn_open() As Boolean

        Dim result As Boolean = False

        Try

            If odbc_conn IsNot Nothing Then
                odbc_conn.Dispose()
                odbc_conn = Nothing
            End If

            conn_ok = False
            odbc_conn = New OdbcConnection '連線資料庫
            odbc_conn.ConnectionString = conn_str
            odbc_conn.Open()
            conn_ok = True
            result = True

        Catch ex As Exception
            conn_ok = False
            '   MsgBox(ex.ToString)
        End Try

        Return result

    End Function

    Public Function conn_close() As Boolean

        Dim result As Boolean = False

        Try
            odbc_conn.Close()
            odbc_conn.Dispose()
            odbc_conn = Nothing
            conn_ok = False
            result = True
        Catch ex As Exception
            conn_ok = False
        End Try

        Return result

    End Function


    Public Function exe_sql(ByVal sql_string As String, ByRef data_ds As DataSet) As Boolean

        Dim result As Boolean = False

        For i As Integer = 0 To cnt

            '  MsgBox(i)

            Try
                If Not result Then
                    If conn_ok = True Then
                        Try
                            Dim odbc_da As New OdbcDataAdapter(sql_string, odbc_conn)
                            odbc_da.Fill(data_ds)
                            odbc_da.Dispose()
                            odbc_da = Nothing
                            result = True
                            Exit For
                        Catch ex As Exception

                            '  MsgBox(ex.ToString)

                        End Try
                    End If

                    If result = False Then
                        Thread.Sleep(1000)
                        conn_close()
                        conn_open()
                    End If
                End If
            Catch ex As Exception
                '   MsgBox(ex.ToString)
            End Try
        Next

        Return result

    End Function
End Class
