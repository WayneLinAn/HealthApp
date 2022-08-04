Imports System.IO
Imports System.Threading
Imports System.IO.Ports
Imports System.Text


Public Class ComUtl

    Private com_start As Boolean = False
    Private str_com_1 As String = "COM1"

    Dim s_com_1 As System.IO.Ports.SerialPort  '設定定位器X COM Port連接端

    Dim parity_1 As System.IO.Ports.Parity '設定定位器X Parity
    Dim baudrate_1 As Integer = 9600   '設定定位器X BaudRate
    Dim data_bit_1 As Integer = 8
    Dim stop_bit_1 As System.IO.Ports.StopBits = Ports.StopBits.One

    Dim conn_com_1 As Boolean = False
    Dim b_com_1_start As Boolean = False

    Private com_thd As Thread


    Public Function StartCom() As Boolean

        '   MsgBox("StartCom Start")

        com_start = True
        b_com_1_start = False

        com_thd = New Thread(AddressOf ReadComOne)
        com_thd.Start()

        ' MsgBox("StartCom End")

        Return True

    End Function


    Public Function StopCom() As Boolean

        com_start = False

        Thread.Sleep(1000)

        If s_com_1 IsNot Nothing Then
            If s_com_1.IsOpen = True Then
                s_com_1.Close()
            End If
            s_com_1.Dispose()
            s_com_1 = Nothing
        End If

        Return True

    End Function


    Private Sub ReadComOne()


        While com_start

            ' 開啟 COM X
            While com_start AndAlso Not b_com_1_start

                Try

                    If s_com_1 IsNot Nothing Then

                        If s_com_1.IsOpen = True Then
                            s_com_1.Close()
                        End If

                        s_com_1.Dispose()
                        s_com_1 = Nothing
                    End If

                    s_com_1 = New SerialPort(str_com_1, 9600, parity_1, data_bit_1, stop_bit_1)

                    s_com_1.Encoding = Encoding.ASCII
                    s_com_1.ReadTimeout = 500

                    s_com_1.Open()
                    conn_com_1 = True
                    b_com_1_start = True

                Catch ex As Exception

                    If s_com_1 IsNot Nothing Then
                        If s_com_1.IsOpen = True Then
                            s_com_1.Close()
                        End If
                        s_com_1.Dispose()
                        s_com_1 = Nothing
                    End If

                Finally
                    Thread.Sleep(1000)
                End Try

            End While

            ' 讀取 COM X
            While com_start AndAlso b_com_1_start

                Try
                    Dim bb(0 To 47) As Byte

                    'Dim tmp_str As String = ""
                    'tmp_str = s_com_1.ReadLine()

                    Dim in_buf_len As Integer = s_com_1.BytesToRead

                    If in_buf_len >= 12 Then

                        Dim len As Integer = s_com_1.Read(bb, 0, 48)
                        Dim tmp_str As String = Encoding.ASCII.GetString(bb)

                        If tmp_str IsNot Nothing AndAlso tmp_str.Length > 0 Then

                            If cur_frm_name = "Mi37" Or cur_frm_name = "Mi58" Then

                                If rs232_mq.Count < 400 Then
                                    rs232_mq.Enqueue(tmp_str)
                                End If

                            End If

                            conn_com_1 = True

                        End If

                    End If

                Catch ex As Exception

                    '   log_mq.Enqueue(ex.ToString)

                End Try

                Thread.Sleep(250)

            End While

            ' log_mq.Enqueue("Com X Has been Disconnected")

            If s_com_1 IsNot Nothing Then
                If s_com_1.IsOpen = True Then
                    s_com_1.Close()
                End If
                s_com_1.Dispose()
                s_com_1 = Nothing
            End If

        End While

    End Sub


End Class
