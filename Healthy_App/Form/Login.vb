Imports System.Text
Imports System.IO

Public Class Login

    Private Sub ReadBasicConf()

        ' Read
        Dim sr As StreamReader
        Dim fn As String

        If Not File.Exists(".\basic_conf.txt") Then
            MsgBox("ERROR!! NO CONF FILE")
        End If

        fn = ".\basic_conf.txt"

        sr = New System.IO.StreamReader(fn, True)

        Dim tmp_str As String
        Dim tmp_str_arr() As String


        While Not sr.EndOfStream

            tmp_str = sr.ReadLine

            If tmp_str.Trim.Length > 0 Then

                tmp_str_arr = tmp_str.Split("=")

                If tmp_str_arr(0) = "srv_ip" Then
                    srv_ip = tmp_str_arr(1)
                End If


            End If

        End While

        sr.Close()
        sr.Dispose()

    End Sub
    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '1. Start Log Mgr
        '2. Start Socket Mgr
        '3. Start DB 
        '4. Start Login

        Try

            If Not logout Then

                ReadBasicConf() '<M1>

                Dim v_ok As Boolean = False

                my_db = New DBUtl
                my_sock = New SockUtl
                my_log = New LogUtl
                my_com = New ComUtl

                If Not my_log.StartLog() Then

                    my_log = Nothing
                    ' Me.Close()

                    '    MsgBox("Log Start Error")

                Else

                    If Not my_db.conn_open() Then

                        my_db = Nothing
                        '  Me.Close()

                        '  MsgBox("DB  Start Error")

                    Else

                        '  MsgBox("DB  Start OK")

                        If Not my_sock.StartSocket() Then

                            my_sock = Nothing
                            '   Me.Close()

                            '    MsgBox("Socket  Start Error")



                        Else

                            ' v_ok = True


                            If Not my_com.StartCom() Then

                                my_com = Nothing

                            Else

                                v_ok = True

                                'tmr_healthy.Enabled = True

                            End If

                        End If


                    End If

                End If

                '####### Temp To do by Kobe####
                'v_ok = True
                '##############################

                If Not v_ok Then

                    txt_emp_id.Enabled = False
                    txt_emp_pwd.Enabled = False
                    btn_login.Enabled = False
                    '   btn_close.Enabled = False
                    'btn_update_pwd.Enabled = False
                    'lbl_error.Visible = True

                End If

            Else

                'tmr_healthy.Enabled = True

            End If



        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
    Private Sub btn_login_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        ExeLogin()
    End Sub

    Private Sub ExeLogin()

        Try

            Dim myds As New DataSet

            Dim sql As String = "Select emp_id, emp_name from emp_auth where emp_id = '" & txt_emp_id.Text.Trim & "' and upper(emp_pwd) = '" & txt_emp_pwd.Text.Trim.ToUpper & "' "

            If my_db.exe_sql(sql, myds) Then

                Dim c As Integer = myds.Tables(0).Rows.Count

                If c > 0 Then

                    user_id = txt_emp_id.Text.Trim

                    If Not TypeOf (myds.Tables(0).Rows(0).Item(1)) Is DBNull Then
                        user_name = myds.Tables(0).Rows(0).Item(1).ToString
                    Else
                        user_name = ""
                    End If

                    'MsgBox("Start")

                    'tmr_healthy.Enabled = False

                    Healthy_App.Mi010.Show()

                    myds = Nothing

                    Me.Hide()



                Else

                    ' Error Handling
                    'my_lab = New LocalAlarmBox
                    'my_lab.lbl_prompt.Text = "帳號密碼錯誤"
                    'my_lab.ShowDialog()
                    MsgBox("error")

                End If

            Else




            End If

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

End Class