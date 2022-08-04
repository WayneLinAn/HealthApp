Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text


Public Class SockUtl

    '  Public srv_ip As String = "192.168.17.247"
    '   Public srv_ip As String = "192.168.17.248"
    Public srv_port As Integer = 50000

    Private Const send_retry As Integer = 1
    Private Const ack_timeout As Integer = 3
    Private Const buf_sz As Integer = 100 * 1024
    Private Const healthy_interval As Integer = 10 ' Rodger@2016.11.30


    Private my_tcp_clnt As TcpClient
    Public conn_ok As Boolean
    Public sock_start As Boolean

    Private send_msg_thd As Thread
    Private rcv_msg_thd As Thread
    Private chk_healthy_thd As Thread  ' Rodger@2016.11.30

    Public healthy_timeout As Boolean = False
    Public heart_beat As Integer = 0  ' Rodger@2016.11.30
    Public error_cnt As Integer = 0  ' Rodger@2016.11.30

    Dim cur As Integer = 0
    Dim work_buf As String = ""

    Public Function StartSocket() As Boolean

        Dim result = False
        sock_start = False

        Try
            Connect()

            If send_msg_thd IsNot Nothing Then
                send_msg_thd.Abort()
                send_msg_thd = Nothing
            End If

            If rcv_msg_thd IsNot Nothing Then
                rcv_msg_thd.Abort()
                rcv_msg_thd = Nothing
            End If

            ' Rodger@2016.11.30
            If chk_healthy_thd IsNot Nothing Then
                chk_healthy_thd.Abort()
                chk_healthy_thd = Nothing
            End If

            sock_start = True

            send_msg_thd = New Thread(AddressOf SendData)
            rcv_msg_thd = New Thread(AddressOf RcvData)
            chk_healthy_thd = New Thread(AddressOf ChkHealthy)  ' Rodger@2016.11.30

            send_msg_thd.Start()
            rcv_msg_thd.Start()
            chk_healthy_thd.Start()  ' Rodger@2016.11.30

            result = True

        Catch ex As Exception
            add_event_log(ex.ToString)
        End Try

        Return result

    End Function

    Public Function StopSocket() As Boolean

        Try
            sock_start = False
            Thread.Sleep(1000)
            DisConnect()
        Catch ex As Exception
            add_event_log(ex.ToString)
        End Try

        Return True

    End Function


    Private Sub Connect()

        conn_ok = False

        add_event_log("Sock starts to Conn")

        Try
            If my_tcp_clnt IsNot Nothing Then
                If conn_ok = True Then
                    my_tcp_clnt.Close()
                End If
                my_tcp_clnt = Nothing
            End If

            my_tcp_clnt = New TcpClient
            my_tcp_clnt.Connect(srv_ip, srv_port)
            my_tcp_clnt.SendTimeout = 0
            my_tcp_clnt.ReceiveTimeout = 5000
            my_tcp_clnt.ReceiveBufferSize = buf_sz
            conn_ok = True

        Catch ex As Exception

            add_event_log(ex.ToString)
            Throw ex

        End Try

        add_event_log("Conn is OK")

    End Sub


    Private Sub DisConnect()

        add_event_log("Sock starts to DisConn")

        Try
            If my_tcp_clnt IsNot Nothing Then
                If conn_ok = True Then
                    my_tcp_clnt.Close()
                End If
                my_tcp_clnt = Nothing
            End If
        Catch ex As Exception
            add_event_log(ex.ToString)
            Throw ex
        End Try

        heart_beat = 0 'Rodger@2016.11.30
        conn_ok = False

        add_event_log("DisConn is Done")

    End Sub


    Private Sub SendData()

        Dim temp_msgdt As MsgDT
        Dim temp_str As String

        While sock_start

            While send_mq.Count > 0

                add_event_log("send_mq.Count [" & send_mq.Count & "]")

                temp_msgdt = send_mq.Dequeue
                temp_str = temp_msgdt.body

                If temp_str IsNot Nothing AndAlso temp_str <> "" Then

                    Dim len As Integer = System.Text.Encoding.UTF8.GetBytes(temp_str).Length
                    Dim [s_byte] As Byte() = System.Text.Encoding.UTF8.GetBytes(temp_str)

                    add_event_log("try to send [" & temp_str & "]")

                    For i As Integer = 1 To send_retry
                        Try
                            If healthy_timeout = True Or conn_ok = False Then
                                DisConnect()
                                Thread.Sleep(1000)   'Rodger@2016.11.30
                                Connect()
                            End If

                            my_tcp_clnt.GetStream.Write([s_byte], 0, len)
                            add_snd_log(temp_str)

                            If CheckAckMsg(temp_msgdt.hdr) Then

                                Exit For
                            End If

                            '       MsgBox("Ack Expire")

                            add_event_log("Ack expire")
                            conn_ok = False

                        Catch ex As Exception
                            conn_ok = False
                            add_event_log(ex.ToString)
                        End Try
                    Next

                End If
            End While

            Thread.Sleep(200)

        End While

    End Sub


    Private Sub RcvData()

        Dim rcv_buffer(buf_sz) As Byte
        Dim rcv_str As String = ""
        Dim rcv_len As Integer

        While sock_start

            Try
                If conn_ok = True AndAlso my_tcp_clnt IsNot Nothing AndAlso my_tcp_clnt.Connected = True Then
                    If my_tcp_clnt.GetStream.CanRead And my_tcp_clnt.GetStream.DataAvailable Then
                        rcv_len = my_tcp_clnt.GetStream.Read(rcv_buffer, 0, buf_sz)
                        'rcv_str = Encoding.UTF8.GetString(rcv_buffer, 0, rcv_len)
                        rcv_str = Encoding.Default.GetString(rcv_buffer, 0, rcv_len)

                        add_event_log(" Sock [" & rcv_len & "][" & rcv_str & "]")

                        ResolveMsg2(rcv_str)

                        heart_beat = 1 'Rodger@2016.11.30 


                    End If
                End If

            Catch ex As Exception
                add_event_log(ex.ToString)
            End Try

            rcv_str = ""

            Thread.Sleep(200)

        End While

    End Sub


    Private Sub ResolveMsg(ByVal rcv_str As String)

        Dim pure_str As String = ""
        Dim temp_str As String = rcv_str
        Dim i, j As Integer
        Dim pure_msgdt As MsgDT
        Dim ack_msgdt As MsgDT
        Dim chr() As String

        While temp_str.Length > 0

            'add_event_log(temp_str)

            '  add_event_log("Rcv Raw Data [" & rcv_str & "]")

            ' MsgBox(temp_str.Length)

            Try
                If Not (temp_str.Contains("~") AndAlso temp_str.Contains("&")) Then
                    Exit While
                End If

                i = temp_str.IndexOf("~")
                j = temp_str.IndexOf("&")

                If Not (j > i) Then
                    Exit While
                End If

                pure_str = temp_str.Substring(i, j - i + 1)

                chr = pure_str.Trim("~").Trim("&").Split("^")

                ' MsgBox(chr(0))

                pure_msgdt.hdr = chr(0)
                pure_msgdt.body = pure_str

                If pure_msgdt.hdr = "9801" Then 'Ack Message
                    ack_msgdt.hdr = chr(5)
                    ack_msgdt.body = ""

                    If ack_mq.Count < 400 Then
                        ack_mq.Enqueue(ack_msgdt)
                    End If

                Else
                    If rcv_mq.Count < 400 Then
                        rcv_mq.Enqueue(pure_msgdt)
                        'add_rcv_log(pure_msgdt.body)
                    End If

                End If

                add_rcv_log(pure_msgdt.body)


                If temp_str.Length > j Then
                    temp_str = temp_str.Substring(j + 1)
                Else
                    Exit While
                End If
            Catch ex As Exception
                add_event_log(ex.ToString)
            End Try

        End While

    End Sub


    Private Sub ResolveMsg2(ByVal rcv_str As String)

        Dim pure_str As String = ""
        Dim temp_str As String = rcv_str
        Dim i, j As Integer
        Dim pure_msgdt As MsgDT
        Dim ack_msgdt As MsgDT
        Dim chr() As String

        If temp_str.Contains("~") Then

            If work_buf.Length <> 0 Then
                work_buf = ""
            End If

            work_buf = temp_str

        Else

            If work_buf.Length = 0 Then
                Exit Sub
            Else
                work_buf &= temp_str
            End If

        End If


        While work_buf.Length > 0

            Try

                If Not (work_buf.Contains("~") AndAlso work_buf.Contains("&")) Then
                    Exit While
                End If

                i = work_buf.IndexOf("~")
                j = work_buf.IndexOf("&")

                If Not (j > i) Then
                    Exit While
                End If

                pure_str = work_buf.Substring(i, j - i + 1)

                chr = pure_str.Trim("~").Trim("&").Split("^")

                pure_msgdt.hdr = chr(0)
                pure_msgdt.body = pure_str

                If pure_msgdt.hdr = "9801" Then 'Ack Message
                    ack_msgdt.hdr = chr(5)
                    ack_msgdt.body = ""

                    If ack_mq.Count < 400 Then
                        ack_mq.Enqueue(ack_msgdt)
                    End If

                Else
                    If rcv_mq.Count < 400 Then
                        rcv_mq.Enqueue(pure_msgdt)

                    End If

                End If

                add_rcv_log(pure_msgdt.body)


                If work_buf.Length > j Then
                    work_buf = work_buf.Substring(j + 1)
                Else
                    Exit While
                End If
            Catch ex As Exception
                add_event_log(ex.ToString)
            End Try

        End While

    End Sub



    Private Function CheckAckMsg(ByVal ack_hdr As String) As Boolean

        Dim result As Boolean = False

        Dim sleep_tm As Integer = 200  '200ms
        Dim cnt As Integer = CInt(ack_timeout * 1000 / sleep_tm)

        Dim temp_msgdt As MsgDT

        Try
            For i As Integer = 1 To cnt

                If ack_mq.Count > 0 Then

                    temp_msgdt = ack_mq.Dequeue

                    If temp_msgdt.hdr <> "" Then
                        If temp_msgdt.hdr = ack_hdr Then
                            result = True
                            ack_mq.Clear()
                            Exit For
                        End If
                    End If
                End If

                Thread.Sleep(sleep_tm)
            Next
        Catch ex As Exception
            add_event_log(ex.ToString)
        End Try

        Return result

    End Function

    'Rodger @ 2016.11.30
    Private Sub ChkHealthy()

        While sock_start

            If conn_ok Then

                If heart_beat = 1 Then

                    error_cnt = 0
                    heart_beat = 0
                    healthy_timeout = False

                Else

                    error_cnt = error_cnt + 1

                    If error_cnt >= 30 Then

                        healthy_timeout = True

                    End If

                End If

            Else

                healthy_timeout = True

            End If

            Thread.Sleep(healthy_interval * 1000)

        End While

    End Sub

End Class



