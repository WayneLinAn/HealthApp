
Imports System.Threading
Imports System.IO


Public Class LogUtl

    Private date_str As String
    Private datetime_str As String

    Private log_thd As Thread

    Private send_msg_log_fn As String
    Private rcv_msg_log_fn As String
    Private event_log_fn As String

    Private send_msg_log_sw As StreamWriter
    Private rcv_msg_log_sw As StreamWriter
    Private event_log_sw As StreamWriter

    Private log_start As Boolean

    Private Const log_keep_day As Integer = 14

    Public Function StartLog() As Boolean

        log_start = False

        date_str = Now.ToString("yyyyMMdd")

        If Not Directory.Exists(".\log") Then
            Directory.CreateDirectory(".\log")
        End If

        send_msg_log_fn = ".\log\send_msg_" & date_str & ".txt"
        rcv_msg_log_fn = ".\log\rcv_msg_" & date_str & ".txt"
        event_log_fn = ".\log\event_" & date_str & ".txt"

        send_msg_log_sw = New System.IO.StreamWriter(send_msg_log_fn, True)
        rcv_msg_log_sw = New System.IO.StreamWriter(rcv_msg_log_fn, True)
        event_log_sw = New System.IO.StreamWriter(event_log_fn, True)

        send_msg_log_sw.AutoFlush = True
        rcv_msg_log_sw.AutoFlush = True
        event_log_sw.AutoFlush = True

        log_start = True

        log_thd = New Thread(AddressOf LogMsg)
        log_thd.Start()

        Return True

    End Function


    Public Function StopLog() As Boolean

        log_start = False

        Thread.Sleep(1000)

        send_msg_log_sw.Close()
        rcv_msg_log_sw.Close()
        event_log_sw.Close()

        Return True

    End Function


    Private Sub LogMsg()

        Dim temp_msgdt As MsgDT

        While log_start

            While log_mq.Count > 0

                CheckLogRotate()

                temp_msgdt = log_mq.Dequeue

                Select Case temp_msgdt.hdr
                    Case "rcv" : LogRcvMsg(temp_msgdt.body)
                    Case "snd" : LogSndMsg(temp_msgdt.body)
                    Case "event" : LogEventMsg(temp_msgdt.body)
                    Case Else 'Error
                End Select

            End While

            Thread.Sleep(100)

        End While

    End Sub

    Private Sub CheckLogRotate()

        datetime_str = Now.ToString("yyyyMMddHHmmss")

        If date_str <> datetime_str.Substring(0, 8) Then
            date_str = datetime_str.Substring(0, 8)

            send_msg_log_sw.Close()
            rcv_msg_log_sw.Close()
            event_log_sw.Close()

            send_msg_log_sw = Nothing
            rcv_msg_log_sw = Nothing
            event_log_sw = Nothing

            send_msg_log_fn = ".\log\send_msg_" & date_str & ".txt"
            rcv_msg_log_fn = ".\log\rcv_msg_" & date_str & ".txt"
            event_log_fn = ".\log\event_" & date_str & ".txt"

            send_msg_log_sw = New System.IO.StreamWriter(send_msg_log_fn, True)
            rcv_msg_log_sw = New System.IO.StreamWriter(rcv_msg_log_fn, True)
            event_log_sw = New System.IO.StreamWriter(event_log_fn, True)

            send_msg_log_sw.AutoFlush = True
            rcv_msg_log_sw.AutoFlush = True
            event_log_sw.AutoFlush = True

            DelOldLogFiles()

        End If

    End Sub


    Private Sub LogRcvMsg(ByVal log_msg As String)
        rcv_msg_log_sw.WriteLine("[" & datetime_str & "] " & log_msg)
        rcv_msg_log_sw.Flush()
    End Sub


    Private Sub LogSndMsg(ByVal log_msg As String)
        send_msg_log_sw.WriteLine("[" & datetime_str & "] " & log_msg)
        send_msg_log_sw.Flush()
    End Sub


    Private Sub LogEventMsg(ByVal log_msg As String)
        event_log_sw.WriteLine("[" & datetime_str & "] " & log_msg)
        event_log_sw.Flush()
    End Sub


    Private Sub DelOldLogFiles()

        Dim dir As String = ".\log"

        If Directory.Exists(dir) Then

            Dim dir_obj As New DirectoryInfo(dir)
            Dim files As FileInfo() = dir_obj.GetFiles("*.*")
            Dim file_name As FileInfo
            Dim last_write_date As DateTime
            Dim temp_ts As TimeSpan

            For Each file_name In files
                Try
                    last_write_date = File.GetCreationTime(file_name.FullName)
                    temp_ts = last_write_date.Subtract(DateTime.Now)
                    If temp_ts.Duration.TotalDays > log_keep_day Then
                        File.Delete(file_name.FullName)
                    End If
                Catch ex As Exception

                End Try
            Next

        End If

    End Sub

   
End Class
