Imports System.Threading
'<M242>
Imports System.Reflection

Module GlobalVar

    Public Const MAX_PASS_COUNT As Short = 17

    Public Const AUTH_NONE As Integer = 0
    Public Const AUTH_READ As Integer = 1
    Public Const AUTH_WRITE As Integer = 2
    Public Const AUTH_SPECIAL As Integer = 3    '<M96>

    Public Const MV_TYPE_SINGLE As Integer = 1
    Public Const MV_TYPE_MULTIPLE As Integer = 2

    Public Const MV_POS_FIRST As Integer = 1
    Public Const MV_POS_LAST As Integer = 2

    Public work_dir As String = Environment.CurrentDirectory
    Public ws As String = Environment.MachineName
    Public user_name As String = "王大明"
    Public user_id As String = "none"
    '<M239>Public ver As String = "10.0"
    '<M242>
    Public ver As Integer = CInt(FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileMajorPart)

    Public cur_ver As String
    Public ver_flag As Boolean

    Public logout As Boolean = False

    Public cur_frm_name As String

    Public srv_ip As String
    Public arch_ip As String = "10.108.137.23"

    Public my_db As DBUtl
    Public my_sock As SockUtl
    Public my_log As LogUtl
    Public my_com As ComUtl

    Public send_mq As New Queue(500)
    Public rcv_mq As New Queue(500)
    Public ack_mq As New Queue(500)
    Public log_mq As New Queue(500)

    Public rcv_alarm_mq As New Queue(500)
    Public rcv_ps_alarm_mq As New Queue(500)

    Public rs232_mq As New Queue(500)

    Public hardcopy_mq As New Queue(500)

    ' Public mtx As New Mutex

    Structure MsgDT
        Public hdr As String
        Public body As String
    End Structure

    Structure ConstDT
        Public item_prog As String
        Public min_v As Double
        Public max_v As Double
    End Structure


    Public open_me As Integer

    Public mv_signal_ref_dn As New Dictionary(Of String, MvSignalRef_T)

    Public alarm_msg_mq As New Queue(500)
    Public alarm_idx As Integer


    Public Const DB_ONLINE As Integer = 0
    Public Const DB_ARCHIVE As Integer = 1

    '--------------------------------------
    'Preset Read
    '--------------------------------------

    Public Const TYPE_NONE As Integer = -1
    Public Const TYPE_ONLINE As Integer = 1
    Public Const TYPE_DEFAULT As Integer = 2
    Public Const TYPE_EMPLOYEE As Integer = 3
    Public Const TYPE_EXCEL As Integer = 4

    Public open_type As Integer = TYPE_NONE
    Public save_type As Integer = TYPE_NONE

    Public Const QUERY_PROD_SC_3 As Integer = 1
    Public Const QUERY_PROD_SC_7 As Integer = 2

    Public query_mode As Integer

    Public prod_sc As String = ""
    Public prod_sc_3 As String = ""
    Public prod_sc_7 As String = ""
    Public prod_size As String = ""

    Public sg_no As String = ""
    Public bb_sno As String = ""
    Public ver_date As String = ""

    Public force_query As Boolean

    '--------------------------------------
    ' For Mi63
    '--------------------------------------

    Public cb_no As Integer = 1

    '--------------------------------------
    ' For Mi64
    '--------------------------------------

    Public pb_no As Integer = 1

    '--------------------------------------
    ' For Mi65
    '--------------------------------------

    Public ib_no As Integer = 1

    '--------------------------------------
    ' For Mi47
    '--------------------------------------

    Public fix_clnt_name As Boolean = False
    Public fix_dest As Boolean = False

    Public fix_dest_str As String = ""
    Public fix_clnt_name_str As String = ""
    Public fix_sg_lname As String = ""

    '--------------------------------------
    ' For Mi25 & Mi99
    '--------------------------------------

    Public odr_lot_no As String = ""

    '--------------------------------------
    'Roll Delay and RHF Too Long
    '--------------------------------------

    Public delay_timestamp As String = ""
    Public delay_min As Integer = 0

    Public stay_rhf_too_long As Integer = 0


    '--------------------------------------
    ' For Mi22, Mi37, Mi58, Mi75, Mi76, Mi77
    '--------------------------------------

    Public mi22_flag As Boolean = False
    Public mi22_disp As Integer = 0
    Public mi22_bb_type As Integer = 0
    Public mi22_lot_no As String = ""
    Public mi22_serial As String = ""
    Public mi22_scs_no As String = ""
    Public mi22_cs_no As String = ""

    '--------------------------------------
    ' For Mi86
    '--------------------------------------

    Public mi86_flag As Boolean = False
    Public mi86_lot_no As String = ""
    Public mi86_serial As String = ""
    Public mi86_scs_no As String = ""
    Public mi86_cs_no As String = ""

    Public mi86_smp_pos As Integer = 0

    '--------------------------------------
    ' For Mi94      '<M404>
    '--------------------------------------

    Public mi94_flag As Boolean = False
    Public mi94_lot_no As String = ""
    Public mi94_serial As String = ""

    Public mi94_pcs_no As String = ""
    Public mi94_pos_no As Integer = 0

    '--------------------------------------
    ' For Mi36
    '--------------------------------------

    Public row_tier_array(0 To 77) As Integer
    Public bundle_array(0 To 38) As Integer

    '<M20>
    'Public mi96_array(0 To 64) As Integer
    'Public mi96_array(0 To 68) As Integer
    '<M303>
    Public mi96_array(0 To 70) As Integer
    Public mi97_array(0 To 75) As Integer

    '--------------------------------------
    ' For Mi17
    '--------------------------------------

    Public mi17_xml_rst As String

    '--------------------------------------
    ' For Mi14
    '--------------------------------------

    Public p_alarm_sfx As Boolean
    Public p_alarm_msg As String = ""

    '--------------------------------------
    ' For Mi37 and Mi58 and Mi50
    '--------------------------------------

    Public mi37_flag As Boolean = False
    Public mi37_smp_shft_no As String = ""

    Public mi37_flat_bb As Integer = 0
    Public mi37_smp_zone As String = ""

    Public mi50_ib_no As Integer = 0 '0:ALL, 1:A1, 2:A2, 3:B

    '<M139>
    Public mi50_db_src As Integer = DB_ONLINE


    '--------------------------------------
    ' For Mi59      '<M345>
    '--------------------------------------

    Public mi59_flag As Boolean = False
    Public mi59_smp_shft_no As String = ""

    Public mi59_flat_bb As Integer = 0





    '--------------------------------------
    ' For Mi12 and Mi14
    '--------------------------------------

    Public FLAT_ALARM_WID_MIN As Integer = 501
    Public FLAT_ALARM_WID_MAX As Integer = 699

    '--------------------------------------
    ' For Mi62 and Mi62_SlabComposition    <M380>
    '--------------------------------------
    Public mi62_heat_no As String = ""
    Public mi62_team_cut_no As String = ""
    ' For Mi62 <M382>
    '--------------------------------------
    Public mi62_DisplayForW411 As Boolean = False
    '--------------------------------------
    ' For Mi31 ADD BY大
    '--------------------------------------
    Public flag_save_alarm As Boolean = False   '<M463>


    Public Sub add_rcv_log(ByVal s As String)

        Dim temp_msgdt As MsgDT

        temp_msgdt.hdr = "rcv"
        temp_msgdt.body = s

        If log_mq.Count < 400 Then
            log_mq.Enqueue(temp_msgdt)
        End If


    End Sub

    Public Sub add_snd_log(ByVal s As String)

        Dim temp_msgdt As MsgDT

        temp_msgdt.hdr = "snd"
        temp_msgdt.body = s

        If log_mq.Count < 400 Then
            log_mq.Enqueue(temp_msgdt)
        End If


    End Sub

    Public Sub add_event_log(ByVal s As String)

        Dim temp_msgdt As MsgDT

        temp_msgdt.hdr = "event"
        temp_msgdt.body = GetMethodInfo() & s

        If log_mq.Count < 400 Then
            log_mq.Enqueue(temp_msgdt)
        End If


    End Sub

    ''' <summary>
    ''' Get Method Information
    ''' </summary>
    ''' <returns>Method Information</returns>
    ''' <remarks></remarks>
    Public Function GetMethodInfo() As String
        Dim ST As New StackTrace(True)
        Return ("[" + ST.GetFrame(2).GetMethod.ReflectedType.Name + "." + ST.GetFrame(2).GetMethod.Name + "()][行:" + ST.GetFrame(2).GetFileLineNumber.ToString + "]")
        'MsgBox(MethodInfo.GetCurrentMethod().ReflectedType.Name + "." + MethodInfo.GetCurrentMethod().Name + "()")
    End Function
    '<M395>僅能輸入英數字
    'ASCII
    '   48~57 : 0~9
    '   65~90 : A~Z
    '   97~122 : a~z
    '   8 : BackSpace鍵
    Public Function NumAlphaLimit(ByVal KeyChar As Char, ByVal txt_v As String) As Char
        Dim KeyAscii As Integer = 0
        Try
            KeyAscii = Asc(KeyChar)
            If KeyAscii <> 8 Then
                If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                    KeyChar = Chr(KeyAscii)
                Else
                    KeyChar = ""
                End If
            End If
        Catch ex As Exception
        End Try
        Return KeyChar
    End Function


    Public Function NumLimit(ByVal KeyChar As Char, ByVal HasDot As Boolean, ByVal txt_v As String) As Char

        Dim KeyAscii As Integer = 0

        Try

            KeyAscii = Asc(KeyChar)

            Dim tmp_dbl As Double

            If KeyAscii <> 8 Then

                If (KeyAscii >= 48 And KeyAscii <= 57) Then


                Else

                    If HasDot Then

                        If KeyAscii = 46 Then

                            If Double.TryParse(txt_v & KeyChar, tmp_dbl) Then
                                KeyChar = Chr(KeyAscii)
                            Else
                                KeyChar = ""
                            End If

                        ElseIf KeyAscii = 45 Then

                            KeyChar = Chr(KeyAscii)

                        Else
                            KeyChar = ""
                        End If

                    Else

                        KeyChar = ""

                    End If

                End If

            End If


        Catch ex As Exception


        End Try


        Return KeyChar

    End Function

    Public Function NumLimitSign(ByVal KeyChar As Char, ByVal HasDot As Boolean, ByVal txt_v As String, ByVal sel_len As Integer, ByVal sel_start As Integer) As Char

        Dim KeyAscii As Integer = 0

        '    MsgBox("sel_start = " & sel_start & ", sel_len = " & sel_len)

        Try

            KeyAscii = Asc(KeyChar)

            Dim tmp_dbl As Double

            If KeyAscii <> 8 Then

                If (KeyAscii >= 48 And KeyAscii <= 57) Then


                Else

                    If HasDot Then

                        If KeyAscii = 46 Then

                            If Double.TryParse(txt_v & KeyChar, tmp_dbl) Then
                                KeyChar = Chr(KeyAscii)
                            Else
                                KeyChar = ""
                            End If

                        ElseIf KeyAscii = 45 Then

                            If sel_len = 0 Then
                                If txt_v.Length = 0 Then
                                    KeyChar = Chr(KeyAscii)
                                ElseIf sel_start = 0 Then
                                    KeyChar = Chr(KeyAscii)
                                Else
                                    KeyChar = ""
                                End If

                            Else
                                KeyChar = Chr(KeyAscii)
                            End If

                        Else
                            KeyChar = ""
                        End If

                    Else

                        KeyChar = ""

                    End If

                End If

            End If


        Catch ex As Exception


        End Try


        Return KeyChar

    End Function


    Public Function NumLimitRange(ByVal KeyChar As Char, ByVal HasDot As Boolean, ByVal txt_v As String, ByVal sel_len As Integer, ByVal sel_start As Integer, ByVal min As Double, ByVal max As Double) As Char

        Dim KeyAscii As Integer = 0

        Try

            KeyAscii = Asc(KeyChar)

            Dim tmp_dbl As Double

            If KeyAscii <> 8 Then

                If (KeyAscii >= 48 And KeyAscii <= 57) Then


                Else

                    If HasDot Then

                        If KeyAscii = 46 Then

                            If Double.TryParse(txt_v & KeyChar, tmp_dbl) Then
                                KeyChar = Chr(KeyAscii)
                            Else
                                KeyChar = ""
                            End If

                        ElseIf KeyAscii = 45 Then

                            'If sel_len = 0 Then

                            '    If txt_v.Length = 0 Then
                            '        KeyChar = Chr(KeyAscii)
                            '    ElseIf sel_start = 0 Then
                            '        KeyChar = Chr(KeyAscii)
                            '    Else
                            '        KeyChar = ""
                            '    End If

                            'Else
                            '    KeyChar = Chr(KeyAscii)
                            'End If

                            'MsgBox("sel_start = " & sel_start & ", sel_len = " & sel_len & ", txt_v = " & txt_v)
                            'MsgBox("-")

                            If sel_start = 0 Or sel_len = txt_v.Length Then
                                KeyChar = Chr(KeyAscii)
                            End If


                        Else
                            KeyChar = ""
                        End If

                    Else

                        KeyChar = ""

                    End If

                End If


                Dim dd As Double = 0
                Dim ss As String = ""

                '      MsgBox("sel_start = " & sel_start & ", sel_len = " & sel_len & ", txt_v = " & txt_v)

                If sel_len = 0 Then
                    '    MsgBox("1")
                    ss = txt_v & KeyChar
                ElseIf sel_len = txt_v.Length Then
                    '   MsgBox("2")
                    ss = KeyChar
                Else
                    '   MsgBox("3")
                    If sel_start = 0 Then
                        '   MsgBox("31")
                        ss = txt_v.Substring(sel_len) & KeyChar
                    Else
                        '  MsgBox("32")
                        ss = txt_v.Substring(0, sel_start) & KeyChar

                        If sel_start + sel_len < txt_v.Length Then
                            '  MsgBox("33")
                            ss &= txt_v.Substring(sel_start + sel_len)
                        End If

                    End If
                End If

                '   MsgBox(ss)

                If Double.TryParse(ss, dd) Then
                    If dd < min Or dd > max Then
                        KeyChar = ""
                    End If
                Else
                    If KeyAscii = 45 Then

                    Else
                        KeyChar = ""
                    End If

                End If

            End If


        Catch ex As Exception


        End Try


        Return KeyChar

    End Function


    Public my_comp_diff As New CompDiffTool

    Public Class CompDiffTool

        Public lst_comp_dt As New List(Of CompDT)

        Public Sub AddDiff(ByVal dt As CompDT)

            add_event_log("AddDiff Start")

            Dim tmp_comp_dt As New CompDT

            tmp_comp_dt.subkey_cnt = dt.subkey_cnt

            For i As Integer = 0 To tmp_comp_dt.subkey_cnt - 1

                tmp_comp_dt.subkey_name(i) = dt.subkey_name(i)
                tmp_comp_dt.subkey_v(i) = dt.subkey_v(i)
            Next

            tmp_comp_dt.item_name = dt.item_name
            tmp_comp_dt.aux_no = dt.aux_no.PadLeft(3, "0")
            tmp_comp_dt.old_v = dt.old_v
            tmp_comp_dt.new_v = dt.new_v

            add_event_log("--------------------------------------------- ")

            add_event_log("AddDiff lst_comp_dt(i).subkey_name(0) = " & tmp_comp_dt.subkey_name(0).ToString)
            add_event_log("AddDiff lst_comp_dt(i).subkey_name(1) = " & tmp_comp_dt.subkey_name(1).ToString)
            add_event_log("AddDiff lst_comp_dt(i).subkey_name(2) = " & tmp_comp_dt.subkey_name(2).ToString)

            add_event_log("AddDiff lst_comp_dt(i).subkey_v(0) = " & tmp_comp_dt.subkey_v(0).ToString)
            add_event_log("AddDiff lst_comp_dt(i).subkey_v(1) = " & tmp_comp_dt.subkey_v(1).ToString)
            add_event_log("AddDiff lst_comp_dt(i).subkey_v(2) = " & tmp_comp_dt.subkey_v(2).ToString)

            add_event_log("--------------------------------------------- ")



            Dim f As Boolean = False
            Dim same_key As Boolean = True

            For i As Integer = 0 To lst_comp_dt.Count - 1

                same_key = True

                If tmp_comp_dt.subkey_cnt > 0 Then

                    If lst_comp_dt(i).subkey_cnt = tmp_comp_dt.subkey_cnt Then

                        For k As Integer = 0 To tmp_comp_dt.subkey_cnt - 1



                            If tmp_comp_dt.subkey_name(k) = lst_comp_dt(i).subkey_name(k) Then


                                If tmp_comp_dt.subkey_v(k) <> lst_comp_dt(i).subkey_v(k) Then

                                    same_key = False
                                    Exit For

                                End If

                            Else

                                same_key = False
                                Exit For

                            End If

                        Next

                    End If

                    If same_key Then

                        add_event_log("AddDiff same_key = true")


                        add_event_log("AddDiff lst_comp_dt(i).subkey_name(0) = " & lst_comp_dt(i).subkey_name(0).ToString)
                        add_event_log("AddDiff lst_comp_dt(i).subkey_name(1) = " & lst_comp_dt(i).subkey_name(1).ToString)
                        add_event_log("AddDiff lst_comp_dt(i).subkey_name(2) = " & lst_comp_dt(i).subkey_name(2).ToString)

                        add_event_log("AddDiff lst_comp_dt(i).subkey_v(0) = " & lst_comp_dt(i).subkey_v(0).ToString)
                        add_event_log("AddDiff lst_comp_dt(i).subkey_v(1) = " & lst_comp_dt(i).subkey_v(1).ToString)
                        add_event_log("AddDiff lst_comp_dt(i).subkey_v(2) = " & lst_comp_dt(i).subkey_v(2).ToString)

                        add_event_log("AddDiff lst_comp_dt(i).item_name & lst_comp_dt(i).aux_no = " & lst_comp_dt(i).item_name & lst_comp_dt(i).aux_no)
                        add_event_log("AddDiff tmp_comp_dt.item_name & tmp_comp_dt.aux_no = " & tmp_comp_dt.item_name & tmp_comp_dt.aux_no)

                        If (lst_comp_dt(i).item_name & lst_comp_dt(i).aux_no) = (tmp_comp_dt.item_name & tmp_comp_dt.aux_no) Then

                            If lst_comp_dt(i).old_v <> tmp_comp_dt.new_v Then
                                lst_comp_dt(i).new_v = tmp_comp_dt.new_v
                            End If


                            add_event_log("old_v =" & lst_comp_dt(i).old_v)
                            add_event_log("new_v =" & lst_comp_dt(i).new_v)

                            add_event_log("AddDiff Update !!")

                            f = True

                            Exit For

                        End If

                    Else

                        add_event_log("AddDiff same_key = false")

                    End If

                Else

                    If (lst_comp_dt(i).item_name & lst_comp_dt(i).aux_no) = (tmp_comp_dt.item_name & tmp_comp_dt.aux_no) Then

                        add_event_log("AddDiff Update !!")

                        If lst_comp_dt(i).old_v <> tmp_comp_dt.new_v Then
                            lst_comp_dt(i).new_v = tmp_comp_dt.new_v
                        End If

                        f = True

                        Exit For

                    End If

                End If

            Next

            If Not f Then

                'If tmp_comp_dt.old_v <> "" AndAlso tmp_comp_dt.old_v <> tmp_comp_dt.new_v Then
                '    lst_comp_dt.Add(tmp_comp_dt)
                '    add_event_log("AddDiff ADD !! ")

                '    ' MsgBox("ADD")

                'End If

                If tmp_comp_dt.old_v <> "" Then
                    If tmp_comp_dt.old_v <> tmp_comp_dt.new_v Then
                        lst_comp_dt.Add(tmp_comp_dt)
                        add_event_log("AddDiff ADD !! ")
                    End If
                Else
                    If tmp_comp_dt.new_v <> "" Then
                        tmp_comp_dt.old_v = tmp_comp_dt.new_v
                        tmp_comp_dt.new_v = ""
                        lst_comp_dt.Add(tmp_comp_dt)
                        add_event_log("AddDiff ADD !! ")
                    End If

                End If


            End If

            add_event_log("AddDiff End")

        End Sub

        Public Sub Clear()
            lst_comp_dt.Clear()
        End Sub

        Public Function MakeMsgData() As String

            Dim rst As String = ""
            Dim cnt As Integer = 0

            For i As Integer = 0 To lst_comp_dt.Count - 1

                If lst_comp_dt(i).new_v <> "" Then

                    rst &= (lst_comp_dt(i).subkey_cnt.ToString & ":")

                    For j As Integer = 0 To lst_comp_dt(i).subkey_cnt - 1
                        rst &= (lst_comp_dt(i).subkey_name(j).PadRight(1) & ":" & lst_comp_dt(i).subkey_v(j).PadRight(1) & ":")
                    Next

                    rst &= (lst_comp_dt(i).item_name.PadRight(1) & ":" & lst_comp_dt(i).aux_no.PadLeft(3, "0") & ":" & lst_comp_dt(i).old_v.PadRight(1) & ":" & lst_comp_dt(i).new_v.PadRight(1) & ":")
                    cnt += 1

                End If

            Next

            Return cnt.ToString & ":" & rst

        End Function

    End Class


    Public Class CompDT

        Public subkey_cnt As Integer

        Public subkey_name(0 To 4) As String
        Public subkey_v(0 To 4) As String

        Public item_name As String

        Public aux_no As String

        Public old_v As String
        Public new_v As String

        Public Sub Clear()

            item_name = ""
            old_v = ""
            new_v = ""
            aux_no = "000"

            subkey_cnt = 0

            For i As Integer = 0 To 4
                subkey_name(i) = ""
                subkey_v(i) = ""
            Next

        End Sub

        Public Sub New()
            item_name = ""
            old_v = ""
            new_v = ""
            aux_no = "000"

            subkey_cnt = 0

            For i As Integer = 0 To 4
                subkey_name(i) = ""
                subkey_v(i) = ""
            Next

        End Sub

    End Class

    Public Class MvSignalRef_T

        Public pos As Integer
        Public mv_type As Integer
        Public zone1 As String
        Public zone2 As String

        Public Sub New(ByVal z1 As String, ByVal z2 As String, ByVal t As Integer, ByVal p As Integer)

            zone1 = z1
            zone2 = z2
            mv_type = t
            pos = p

        End Sub

    End Class

    '<M244>
    Public Function chkOutofRange(ByRef valNew As Double, ByRef valOld As Double, ByVal threshold As Double, ByVal chkoption As Boolean) As Boolean
        '比較新值valNew及舊值valOld之值是否超過threshold範圍，option為true時比較實際值，否則比較舊值的變更比例
        '回傳true表示在合法範圍內，否則回傳false
        Select Case chkoption
            Case True
                If Math.Round(Math.Abs(valNew - valOld), 2) >= threshold Then Return False Else Return True
            Case False
                If Math.Round(Math.Abs(((valNew - valOld) / valOld)), 2) >= threshold Then Return False Else Return True
            Case Else
                Return False
        End Select
    End Function

    '<M463>    '<M506> 原先只有Mi031在用   已搬到Mi031 上 並改名 chkSGPos_Mi031
    Public Function chkSGPos(ByRef CB_Pos As Double, ByRef DS_Front As Double, ByRef WS_Front As Double, ByVal DS_Behind As Double, ByRef WS_Behind As Double, ByRef flag As Integer, ByVal alarm_index_arr As Array)
        '比較孔槽POS及SG的POS
        If CB_Pos - DS_Behind < 0 Then
            flag_save_alarm = True
            alarm_index_arr(2) = True
        End If
        If CB_Pos - DS_Front < 0 Then
            flag_save_alarm = True
            alarm_index_arr(0) = True
        End If

        If CB_Pos + WS_Behind > 2800 Then
            flag_save_alarm = True
            alarm_index_arr(3) = True
        End If
        If CB_Pos + WS_Front > 2800 Then
            flag_save_alarm = True
            alarm_index_arr(1) = True
        End If

        If ((flag + 1) Mod 2) = 0 Then
            If (DS_Front + WS_Front) < (DS_Behind + WS_Behind) Then
                alarm_index_arr(0) = True
                alarm_index_arr(1) = True
                alarm_index_arr(2) = True
                alarm_index_arr(3) = True
                'flag_save_alarm = True    '<M506>  W411祐誠 要放寬限制  所以這行要註解 
            End If
        Else
            If (DS_Front + WS_Front) > (DS_Behind + WS_Behind) Then
                alarm_index_arr(0) = True
                alarm_index_arr(1) = True
                alarm_index_arr(2) = True
                alarm_index_arr(3) = True
                'flag_save_alarm = True    '<M506>  W411祐誠 要放寬限制  所以這行要註解 
            End If
        End If
        'flag_save_alarm = False

        Return True
    End Function

End Module
