Option Strict Off
Option Explicit On
Imports System.Text
Imports VB = Microsoft.VisualBasic
Imports System.IO
Imports System.Data.Odbc
Imports System.Runtime.InteropServices
Public Class monitoring
    'database
    Dim jumlah_part(8) As Integer
    'history
    Dim node_history(8, 200) As Integer
    Public jum_part(8, 200) As Integer
    'record liffting
    Public lif1, lif2, lif3, lif4, lif5, lif6, lif7, lif8 As Integer

    'waktu
    Dim waktu As Integer
    '/waktu
    Dim SndBuf(1024) As Byte
    Dim RcvBuf(1024) As Byte
    Dim GwNode As Short
    Dim GwMax As Short
    Dim GwNodePort As Short
    Dim glblConnectStart As Boolean
    Dim glbnConnectCount As Short
    Dim rcv_ccb As QWAY_CCB
    Dim lAutoScan As Boolean
    Dim dispnum As Integer
    Dim lInitClearTags As Boolean
    Dim nShowReceiveSta As Integer
    'fungsi untuk write file .ini
    Public Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringW" (ByVal lpSection As String, ByVal lpParamName As String,
    ByVal lpParamVal As String, ByVal lpFileName As String) As Int32
    Public Sub WriteErrorx(ByVal Error_Messagex As String)
        Dim objStreamWrite As StreamWriter
        objStreamWrite = New StreamWriter("information-log.txt", True, Encoding.Unicode)
        objStreamWrite.WriteLine(Error_Messagex)
        objStreamWrite.Close()
    End Sub
    'procedure untuk write .ini
    Public Sub writeini(ByVal iniFilename As String, ByVal section As String, ByVal ParamName As String, ByVal ParamVal As String)
        'menanggil fungsi WritePrivateProfilString untuk write file .ini
        Dim result As Integer = WritePrivateProfileString(section, ParamName, ParamVal, iniFilename)
    End Sub
    'function untuk read file .ini
    Public Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" (ByVal lpSection As String, ByVal lpParamName As String,
    ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32,
    ByVal lpFilename As String) As Int32

    'function untuk read file .ini
    Public Function readini(ByVal iniFileName As String, ByVal Section As String, ByVal ParamName As String, ByVal ParamDefault As String) As String
        Dim ParamVal As String = Space$(1024)
        Dim LenParamVal As Long = GetPrivateProfileString(Section, ParamName, ParamDefault, ParamVal, Len(ParamVal), iniFileName)
        'mengembalikan nilai yang sudah didapatkan
        readini = Strings.Left(ParamVal, LenParamVal)
    End Function
    Public Sub liffting_default()
        TextBox4.Text = readini(FilePath, "HISTORY", "pos1", "")
        TextBox5.Text = readini(FilePath, "HISTORY", "pos2", "")
        TextBox9.Text = readini(FilePath, "HISTORY", "pos3", "")
        TextBox13.Text = readini(FilePath, "HISTORY", "pos4", "")
        TextBox17.Text = readini(FilePath, "HISTORY", "pos5", "")
        TextBox21.Text = readini(FilePath, "HISTORY", "pos6", "")
        TextBox25.Text = readini(FilePath, "HISTORY", "pos7", "")
        TextBox29.Text = readini(FilePath, "HISTORY", "pos8", "")
        lif1 = readini(FilePath, "HISTORY", "pos1", "")
        lif2 = readini(FilePath, "HISTORY", "pos2", "")
        lif3 = readini(FilePath, "HISTORY", "pos3", "")
        lif4 = readini(FilePath, "HISTORY", "pos4", "")
        lif5 = readini(FilePath, "HISTORY", "pos5", "")
        lif6 = readini(FilePath, "HISTORY", "pos6", "")
        lif7 = readini(FilePath, "HISTORY", "pos7", "")
        lif8 = readini(FilePath, "HISTORY", "pos8", "")
    End Sub
    Public Sub write_pos1()
        writeini(FilePath, "HISTORY", "pos1", TextBox4.Text)
    End Sub
    Public Sub write_pos2()
        writeini(FilePath, "HISTORY", "pos2", TextBox5.Text)
    End Sub
    Public Sub write_pos3()
        writeini(FilePath, "HISTORY", "pos3", TextBox9.Text)
    End Sub
    Public Sub write_pos4()
        writeini(FilePath, "HISTORY", "pos4", TextBox13.Text)
    End Sub
    Public Sub write_pos5()
        writeini(FilePath, "HISTORY", "pos5", TextBox17.Text)
    End Sub
    Public Sub write_pos6()
        writeini(FilePath, "HISTORY", "pos6", TextBox21.Text)
    End Sub
    Public Sub write_pos7()
        writeini(FilePath, "HISTORY", "pos7", TextBox25.Text)
    End Sub
    Public Sub write_pos8()
        writeini(FilePath, "HISTORY", "pos8", TextBox29.Text)
    End Sub

    Sub andon_count(node_part, jum_part)
        Dim actual, min, kritis, hasil As Integer
        Dim status As String
        strin = "select *from data_part where node = '" & CInt(node_part) & "'"
        cmd = New OdbcCommand(strin, connection_andon)
        rd = cmd.ExecuteReader()
        While rd.Read()
            actual = CInt(rd("QTY ACTUAL"))
            min = CInt(rd("BATAS MINIMAL"))
            kritis = CInt(rd("BATAS KRITIS"))
        End While
        hasil = (actual - jum_part)
        status = ""
        If hasil <= min And hasil > kritis Then
            status = "2"
        ElseIf hasil <= kritis Then
            status = "1"
        End If
        Dim update_part As String = "UPDATE `data_part` SET `QTY ACTUAL` = '" & hasil & "' , `STATUS` = '" & status & "' WHERE node = '" & CInt(node_part) & "'"
        cmd = New OdbcCommand(update_part, connection_andon)
        cmd.ExecuteNonQuery()
    End Sub
    Sub sub_panggil(pos)
        Try
            Dim nod As Integer = 0
            Dim a As Integer
            Dim node As Integer
            Dim ret As Integer, nLamp As Byte
            Module3.numvalue(pos) = 0
            strin = "select *from tb_master_pick_sps_final where pos = '" & pos & "' and suffix = '" & menu_form.suffix1(pos) & "' and type = '" & menu_form.model1(pos) & "' and color_code = '" & menu_form.color1(pos) & "' and jum_part <> '0'"
            cmd = New OdbcCommand(strin, connection)
            rd = cmd.ExecuteReader()
            While rd.Read()
                Call Module3.tambahvalue(pos, 1)
                jumlah_part(pos) += 1
                node_history(pos, jumlah_part(pos)) = rd("node")
                jum_part(pos, jumlah_part(pos)) = rd("Jum_part")
                nLamp = Val(Mid(1, 1, 2))
                ret = AB_LB_DspAddr(CurGwId, rd("node"))
                ret = AB_LED_Status(CurGwId, rd("node"), 1, nLamp)
                ret = AB_LB_DspNum(CurGwId, rd("node"), rd("Jum_part"), 2, 0)
                node = rd("node")
                menu_form.part_name(node) = pos & " | " & rd("node") & " | " & rd("part_name")
                If pos = 1 Then
                    ListBox1.Items.Add(menu_form.part_name(node))
                ElseIf pos = 2 Then
                    ListBox2.Items.Add(menu_form.part_name(node))
                ElseIf pos = 3 Then
                    ListBox3.Items.Add(menu_form.part_name(node))
                ElseIf pos = 4 Then
                    ListBox4.Items.Add(menu_form.part_name(node))
                ElseIf pos = 5 Then
                    ListBox5.Items.Add(menu_form.part_name(node))
                ElseIf pos = 6 Then
                    ListBox6.Items.Add(menu_form.part_name(node))
                ElseIf pos = 7 Then
                    ListBox7.Items.Add(menu_form.part_name(node))
                ElseIf pos = 8 Then
                    ListBox8.Items.Add(menu_form.part_name(node))
                End If
                ' MsgBox("ada")
                Call Module3.value(pos)
                ' ret = AB_LED_Dsp(CurGwId, rd("node"), nLamp, nLed)
            End While
            If Module3.numvalue(pos) = 0 Then
                'MsgBox("ga ada")
                Dim nodesample As Integer
                If pos = 1 Then
                    'confirmasi
                    nodesample = 180
                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                    ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                    'lampu view pos

                    ' Dim node As Integer, ret As Integer, value As String, nDot As Byte, nLed As Integer
                    ' node = 181
                    ' value = "Then Thenabcdef"
                    ' nLed = GetLEDStatus(Trim("0:Normal display"))
                    ' nDot = DotConvert(CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, CheckBox7.Checked, CheckBox8.Checked)
                    ' ret = AB_LB_DspStr(CurGwId, node, value, nDot, nLed)

                ElseIf pos = 2 Then
                    nodesample = 182
                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                    ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                ElseIf pos = 3 Then
                    nodesample = 184
                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                    ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                ElseIf pos = 4 Then
                    nodesample = 186
                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                    ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                ElseIf pos = 5 Then
                    nodesample = 188
                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                    ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                ElseIf pos = 6 Then
                    nodesample = 190
                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                    ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                ElseIf pos = 7 Then
                    nodesample = 192
                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                    ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                ElseIf pos = 8 Then
                    nodesample = 194
                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                    ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                    'confirm
                End If
            End If
            menu_form.jum_part1 = a
        Catch ex As Exception
            If pos = 1 Then
                Dim numerror As Integer = CInt(log1.Text)
                numerror += 1
                log1.Text = numerror
            ElseIf pos = 2 Then
                Dim numerror As Integer = CInt(log2.Text)
                numerror += 1
                log2.Text = numerror
            ElseIf pos = 3 Then
                Dim numerror As Integer = CInt(log3.Text)
                numerror += 1
                log3.Text = numerror
            ElseIf pos = 4 Then
                Dim numerror As Integer = CInt(log4.Text)
                numerror += 1
                log4.Text = numerror
            ElseIf pos = 5 Then
                Dim numerror As Integer = CInt(log5.Text)
                numerror += 1
                log5.Text = numerror
            ElseIf pos = 6 Then
                Dim numerror As Integer = CInt(log6.Text)
                numerror += 1
                log6.Text = numerror
            ElseIf pos = 7 Then
                Dim numerror As Integer = CInt(log7.Text)
                numerror += 1
                log7.Text = numerror
            ElseIf pos = 8 Then
                Dim numerror As Integer = CInt(log8.Text)
                numerror += 1
                log8.Text = numerror
            End If
            WriteErrorx("[GET DATA GAGAL]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "- GET DATA GAGAL DARI POS " & pos & " DARI TABLE tb_master_pick_sps_final")
        End Try
    End Sub
    Sub bersih()
        Dim a As Integer
        For a = 1 To 179
            menu_form.part_name(a) = 0
        Next

        For a = 1 To 8
            numvalue(a) = 0
        Next
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox5.Items.Clear()
        ListBox6.Items.Clear()
        ListBox7.Items.Clear()
        ListBox8.Items.Clear()
    End Sub
    Sub belum_ada(pos_sample)
        If pos_sample = "1" Then
            TextBox1.Text = " - "
            TextBox2.Text = " - "
            TextBox3.Text = " - "
        ElseIf pos_sample = "2" Then
            TextBox8.Text = " - "
            TextBox7.Text = " - "
            TextBox6.Text = " - "
        ElseIf pos_sample = "3" Then
            TextBox12.Text = " - "
            TextBox11.Text = " - "
            TextBox10.Text = " - "
        ElseIf pos_sample = "4" Then
            TextBox16.Text = " - "
            TextBox15.Text = " - "
            TextBox14.Text = " - "
        ElseIf pos_sample = "5" Then
            TextBox20.Text = " - "
            TextBox19.Text = " - "
            TextBox18.Text = " - "
        ElseIf pos_sample = "6" Then
            TextBox24.Text = " - "
            TextBox23.Text = " - "
            TextBox22.Text = " - "
        ElseIf pos_sample = "7" Then
            TextBox28.Text = " - "
            TextBox27.Text = " - "
            TextBox26.Text = " - "
        ElseIf pos_sample = "8" Then
            TextBox32.Text = " - "
            TextBox31.Text = " - "
            TextBox30.Text = " - "
        End If
    End Sub
    Sub view_liffting(node1, suffix, liffting)
        Dim node As Integer, ret As Integer, value As String, nLed As Integer
        node = CInt(node1)
        value = Trim(suffix & "" & liffting)
        nLed = GetLEDStatus(Trim(0))
        ret = AB_LB_DspStr(CurGwId, node, value, 0, nLed)
    End Sub
    Sub save_history(pos)
        koneksi_andon()
        If koneksi_status_andon = True Then
            Try
                cmd = connection.CreateCommand
                cmd.CommandText = "insert into tracert_sps_final value('POS " & pos & "','" & menu_form.liffting1(pos) & "' , '" & menu_form.vin1(pos) & "' , '" & menu_form.suffix1(pos) & "' , '" & menu_form.model1(pos) & "', '" & menu_form.color1(pos) & "' , '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "' , 'OK' , '-')"
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                If pos = 1 Then
                    Dim numerror As Integer = CInt(log1.Text)
                    numerror += 1
                    log1.Text = numerror
                ElseIf pos = 2 Then
                    Dim numerror As Integer = CInt(log2.Text)
                    numerror += 1
                    log2.Text = numerror
                ElseIf pos = 3 Then
                    Dim numerror As Integer = CInt(log3.Text)
                    numerror += 1
                    log3.Text = numerror
                ElseIf pos = 4 Then
                    Dim numerror As Integer = CInt(log4.Text)
                    numerror += 1
                    log4.Text = numerror
                ElseIf pos = 5 Then
                    Dim numerror As Integer = CInt(log5.Text)
                    numerror += 1
                    log5.Text = numerror
                ElseIf pos = 6 Then
                    Dim numerror As Integer = CInt(log6.Text)
                    numerror += 1
                    log6.Text = numerror
                ElseIf pos = 7 Then
                    Dim numerror As Integer = CInt(log7.Text)
                    numerror += 1
                    log7.Text = numerror
                ElseIf pos = 8 Then
                    Dim numerror As Integer = CInt(log8.Text)
                    numerror += 1
                    log8.Text = numerror
                End If
                WriteErrorx("[SAVE HISTORY GAGAL]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "- Save Data To TB tracert_sps_final Failed")
            End Try
        Else
            WriteErrorx("[SAVE HISTORY GAGAL]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "- Cannot Connect to TB Tracert_sps_final")
        End If
    End Sub
    Sub panggil(pos, lifft)
        koneksi()
        koneksi_scan()
        If koneksi_status_scan = True Then
            Try
                Dim nod As Integer
                Dim index As Integer
                Dim ret As Integer, nLamp As Byte
                If jumlah_part(pos) = 0 Then
                Else
                    For index = 1 To jumlah_part(pos)
                        nod = nod + 1
                        nLamp = Val(Mid(1, 1, 1))
                        ret = AB_LB_DspAddr(CurGwId, node_history(pos, nod))
                        ret = AB_LED_Status(CurGwId, node_history(pos, nod), 0, nLamp)
                        ret = AB_LB_DspNum(CurGwId, node_history(pos, nod), 0, 0, 0)
                        node_history(pos, node_history(pos, nod)) = 0
                    Next
                    jumlah_part(pos) = 0
                End If

                ' Dim str_liffting As String
                menu_form.liffting1(pos) = ""
                menu_form.suffix1(pos) = ""
                menu_form.vin1(pos) = ""
                menu_form.model1(pos) = ""
                menu_form.color1(pos) = ""
                'connection.Close()
                Dim pos_sample As Integer = pos
                If pos_sample = "1" Then
                    ListBox1.Items.Clear()
                    belum_ada(pos_sample)
                    strin = "SELECT * FROM scan where Liffting = " & CInt(TextBox4.Text)
                ElseIf pos_sample = "2" Then
                    ListBox2.Items.Clear()
                    belum_ada(pos_sample)
                    strin = "SELECT * FROM scan where Liffting = " & CInt(TextBox5.Text)
                ElseIf pos_sample = "3" Then
                    ListBox3.Items.Clear()
                    belum_ada(pos_sample)
                    strin = "SELECT * FROM scan where Liffting = " & CInt(TextBox9.Text)
                ElseIf pos_sample = "4" Then
                    ListBox4.Items.Clear()
                    belum_ada(pos_sample)
                    strin = "SELECT * FROM scan where Liffting = " & CInt(TextBox13.Text)
                ElseIf pos_sample = "5" Then
                    ListBox5.Items.Clear()
                    belum_ada(pos_sample)
                    strin = "SELECT * FROM scan where Liffting = " & CInt(TextBox17.Text)
                ElseIf pos_sample = "6" Then
                    ListBox6.Items.Clear()
                    belum_ada(pos_sample)
                    strin = "SELECT * FROM scan where Liffting =" & CInt(TextBox21.Text)
                ElseIf pos_sample = "7" Then
                    ListBox7.Items.Clear()
                    belum_ada(pos_sample)
                    strin = "SELECT * FROM scan where Liffting = " & CInt(TextBox25.Text)
                ElseIf pos_sample = "8" Then
                    ListBox8.Items.Clear()
                    belum_ada(pos_sample)
                    strin = "SELECT * FROM scan where Liffting = " & CInt(TextBox29.Text)
                End If
                cmd = New OdbcCommand(strin, connection_scan)
                rd2 = cmd.ExecuteReader
                While rd2.Read
                    menu_form.liffting1(pos_sample) = rd2("liffting")
                    menu_form.suffix1(pos_sample) = rd2("suffix")
                    menu_form.vin1(pos_sample) = rd2("vin")
                    menu_form.model1(pos_sample) = rd2("model_code")
                    menu_form.color1(pos_sample) = rd2("color_code")
                    menu_form.entry(pos_sample) = rd2("Entry_Date")
                End While
                If pos_sample = "1" Then
                    Dim conver As String = CInt(TextBox4.Text)
                    If conver < 9999 Then
                        conver = conver + 1
                    Else
                        conver = 1
                    End If
                    TextBox33.Text = conver
                    TextBox1.Text = menu_form.vin1(1)
                    TextBox2.Text = menu_form.suffix1(1)
                    TextBox3.Text = menu_form.model1(1)
                    Call sub_panggil(1)
                    Call view_liffting(181, menu_form.suffix1(1), menu_form.liffting1(1))
                ElseIf pos_sample = "2" Then
                    TextBox8.Text = menu_form.vin1(2)
                    TextBox7.Text = menu_form.suffix1(2)
                    TextBox6.Text = menu_form.model1(2)
                    Call sub_panggil(2)
                    Call view_liffting(183, menu_form.suffix1(2), menu_form.liffting1(2))
                ElseIf pos_sample = "3" Then
                    TextBox12.Text = menu_form.vin1(3)
                    TextBox11.Text = menu_form.suffix1(3)
                    TextBox10.Text = menu_form.model1(3)
                    Call sub_panggil(3)
                    Call view_liffting(185, menu_form.suffix1(3), menu_form.liffting1(3))
                ElseIf pos_sample = "4" Then
                    TextBox16.Text = menu_form.vin1(4)
                    TextBox15.Text = menu_form.suffix1(4)
                    TextBox14.Text = menu_form.model1(4)
                    Call sub_panggil(4)
                    Call view_liffting(187, menu_form.liffting1(4), menu_form.suffix1(4))
                    'Call view_liffting(187, "_", menu_form.liffting1(4))
                ElseIf pos_sample = "5" Then
                    TextBox20.Text = menu_form.vin1(5)
                    TextBox19.Text = menu_form.suffix1(5)
                    TextBox18.Text = menu_form.model1(5)
                    Call sub_panggil(5)
                    Call view_liffting(189, menu_form.suffix1(5), menu_form.liffting1(5))
                    'Call view_liffting(189, "_", menu_form.liffting1(5))
                ElseIf pos_sample = "6" Then
                    TextBox24.Text = menu_form.vin1(6)
                    TextBox23.Text = menu_form.suffix1(6)
                    TextBox22.Text = menu_form.model1(6)
                    Call sub_panggil(6)
                    Call view_liffting(191, menu_form.suffix1(6), menu_form.liffting1(6))
                    Call view_liffting(191, "_", menu_form.liffting1(6))
                ElseIf pos_sample = "7" Then
                    TextBox28.Text = menu_form.vin1(7)
                    TextBox27.Text = menu_form.suffix1(7)
                    TextBox26.Text = menu_form.model1(7)
                    Call sub_panggil(7)
                    'Call view_liffting(193, "_", menu_form.liffting1(7))
                    Call view_liffting(193, menu_form.suffix1(7), menu_form.liffting1(7))
                ElseIf pos_sample = "8" Then
                    TextBox32.Text = menu_form.vin1(8)
                    TextBox31.Text = menu_form.suffix1(8)
                    TextBox30.Text = menu_form.model1(8)
                    Call sub_panggil(8)
                    'Call view_liffting(195, "_", menu_form.liffting1(8))
                    Call view_liffting(195, menu_form.suffix1(8), menu_form.liffting1(8))
                End If
            Catch ex As Exception
                MsgBox("GET DATA GAGAL DARI TABLE SCAN | CEK JARINGAN ANDA")
                WriteErrorx("[GET DATA FAILED FROM TABLE SCAN / CONTROLLER NOT CONNECT]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "-" & ex.Message)
            End Try
        Else
            If pos = 1 Then
                Dim numerror As Integer = CInt(log1.Text)
                numerror += 1
                log1.Text = numerror
            ElseIf pos = 2 Then
                Dim numerror As Integer = CInt(log2.Text)
                numerror += 1
                log2.Text = numerror
            ElseIf pos = 3 Then
                Dim numerror As Integer = CInt(log3.Text)
                numerror += 1
                log3.Text = numerror
            ElseIf pos = 4 Then
                Dim numerror As Integer = CInt(log4.Text)
                numerror += 1
                log4.Text = numerror
            ElseIf pos = 5 Then
                Dim numerror As Integer = CInt(log5.Text)
                numerror += 1
                log5.Text = numerror
            ElseIf pos = 6 Then
                Dim numerror As Integer = CInt(log6.Text)
                numerror += 1
                log6.Text = numerror
            ElseIf pos = 7 Then
                Dim numerror As Integer = CInt(log7.Text)
                numerror += 1
                log7.Text = numerror
            ElseIf pos = 8 Then
                Dim numerror As Integer = CInt(log8.Text)
                numerror += 1
                log8.text = numerror
            End If
            WriteErrorx("[GET DATA GAGAL]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "- GET DATA GAGAL DARI POS " & pos & " di LIFFTING " & lifft & " KARENA JARINGAN ERROR")
            MsgBox("GET DATA GAGAL DI POS " & pos & " DI LIFFTING " & lifft & "| CEK JARINGAN ANDA")
        End If
    End Sub
    ' ret = AB_LED_Status(CurGwId, rd("node"), 3, nLamp)
    Function GetLEDStatus(ByVal cLED As String) As Integer
        Dim nValue As Short
        If Val(cLED) > 0 Then
            nValue = Val(cLED)
        Else
            If Val(Mid(cLED, 1, 2)) < 0 Then
                nValue = Val(Mid(cLED, 1, 2))
            Else
                If Mid(cLED, 1, 1) = ">" Then    'default sample data
                    nValue = 300
                End If
            End If
        End If
        GetLEDStatus = nValue
    End Function
    Sub keadaan_on(pos)
        Dim node(100) As Integer
        Dim angka As Integer = 0
        Dim repet As Integer = 200
        Dim repetan As Integer = 0
        Dim enda As Integer = 2
        'Module3.numvalue = 0
        strin = "select *from tb_master_pick_sps_final where pos = '" & pos & "'"
        cmd = New OdbcCommand(strin, connection)
        rd = cmd.ExecuteReader()
        While rd.Read()
            For repetan = 1 To enda
                If Not node((repetan) = rd("node")) And repetan = enda Then
                    node(enda) = rd("node")
                    enda = enda + 1
                Else
                    Return
                End If
            Next
        End While
        While enda
            Dim nLamp As Byte = Val(Mid(1, 1, 1))
            Dim nColor As Byte = Val(Mid(0, 1, 1))
            Dim ret As Byte = AB_LED_Status(CurGwId, node(enda), 5, nLamp)
            node(enda) = 0
        End While
        enda = 2
        If pos = 1 Then
            ListBox1.Items.Clear()
        ElseIf pos = 2 Then
            ListBox2.Items.Clear()
        ElseIf pos = 3 Then
            ListBox3.Items.Clear()
        ElseIf pos = 4 Then
            ListBox4.Items.Clear()
        ElseIf pos = 5 Then
            ListBox5.Items.Clear()
        ElseIf pos = 6 Then
            ListBox6.Items.Clear()
        ElseIf pos = 7 Then
            ListBox7.Items.Clear()
        ElseIf pos = 8 Then
            ListBox8.Items.Clear()
        End If
    End Sub
    Sub cek(n)
        Dim cek_n As Integer
        Dim node(100) As Integer
        For cek_n = 1 To 200
            If Not node(cek_n) = n Then

            End If
        Next
    End Sub
    Sub keadaan_off(pos)
        Dim node(200) As Integer
        Dim angka As Integer = 0
        Dim repet As Integer = 200
        Dim enda As Integer = 2
        Dim repetan As Integer = 0
        'Module3.numvalue = 0
        strin = "select *from tb_master_pick_sps_final where POS = '" & pos & "'"
        cmd = New OdbcCommand(strin, connection)
        rd = cmd.ExecuteReader()
        While rd.Read()
            For repetan = 1 To enda
                If Not node((repetan) = rd("node")) And repetan = enda Then
                    node(enda) = rd("node")
                    enda = enda + 1
                Else
                    Return
                End If
            Next
        End While
        While enda
            Dim nLamp As Byte = Val(Mid(1, 1, 1))
            Dim nColor As Byte = Val(Mid(0, 1, 1))
            Dim ret As Byte = AB_LED_Status(CurGwId, node(enda), 0, nLamp)
            node(enda) = 0
        End While
        enda = 2
        If pos = 1 Then
            ListBox1.Items.Clear()
        ElseIf pos = 2 Then
            ListBox2.Items.Clear()
        ElseIf pos = 3 Then
            ListBox3.Items.Clear()
        ElseIf pos = 4 Then
            ListBox4.Items.Clear()
        ElseIf pos = 5 Then
            ListBox5.Items.Clear()
        ElseIf pos = 6 Then
            ListBox6.Items.Clear()
        ElseIf pos = 7 Then
            ListBox7.Items.Clear()
        ElseIf pos = 8 Then
            ListBox8.Items.Clear()
        End If
    End Sub
    Private Sub Dap_Setup()
        Dim ret As Short, nGw As Integer
        If CurGwId = 0 Then
            For nGw = 1 To GwCnt
                ret = AB_GW_Open(nGw) 'Open the specified Gateway
            Next nGw
        Else
            ret = AB_GW_Open(CurGwId) 'Open the specified Gateway
        End If
    End Sub
    Private Sub monitoring_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
        liffting_default()
        Try
            Dim ret As Short
            ChDrive(Mid(My.Application.Info.DirectoryPath, 1, 1))
            ChDir(My.Application.Info.DirectoryPath)
            ret = AB_API_Open() 'Startup and initialize the ABLELINK API
            LoadInitial()
            DelayItem = ""
            lAutoScan = True
            lSetBestPoll = False
            lInitClearTags = False
            dispnum = 1
            ScanTimer = 100         'generally value = 100
            Timer1.Interval = ScanTimer
            CurGwId = 0   'Open All Gateway (Below in Dap_Setup())
            Dap_Setup()
            'Default 1st Gateway 
            CurGwId = 1
            'ret = AB_LED_Status(CurGwId, 252, 1, 1)
        Catch ex As Exception
            MsgBox("connection failed to Controller")
        End Try
    End Sub
    Private Sub InitListBox1()
        Dim nGw As Integer
        If Form1.ListBox1.Items.Count = 0 Then
            For nGw = 1 To GwCnt
                Form1.ListBox1.Items.Add("Gw" & nGw & ":" & GwConf(nGw - 1).ip)
            Next nGw
        End If
    End Sub
    Function InitClearTags(ByVal nGw As Integer) As Integer
        Dim ret As Integer
        If lInitClearTags = False Then
            ret = AB_LB_DspStr(nGw, -252, 0, 0, -3)
            lInitClearTags = True
        End If
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim node As Integer, ret As Integer
            node = CInt(252)
            ret = AB_LB_DspNum(CurGwId, node, 1, 1, -3)
            ret = AB_AHA_ClrDsp(CurGwId, node)
            Call bersih()
        Catch ex As Exception
            WriteErrorx("[controller proglem]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "-" & ex.Message)
            MsgBox("KONEKSI KE CONTROLLER GAGAL | CEK KABEL JARINGAN ANDA")
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Call keadaan_on(1)
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs)
        keadaan_off(1)
    End Sub
    Private Sub BlinkAutoStatus()
        If lAutoScan = True Then
            If nShowReceiveSta = 6 Then
                If Form1.GroupBox5.Text = "¡³" & " Receive Data" Then
                    Form1.GroupBox5.Text = "¡´" & " Receive Data"
                Else
                    Form1.GroupBox5.Text = "¡³" & " Receive Data"
                End If
                nShowReceiveSta = 0
            End If
            nShowReceiveSta = nShowReceiveSta + 1
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim nGw As Integer, ret As Integer, nNum As Integer
        nNum = 0
        Timer1.Enabled = False        'connct gateway
        InitListBox1()
        For nGw = 1 To GwCnt
            ret = AB_GW_Status(nGw)                        'Get the status of current Gateway or specified Gateway
            If lAutoScan = True Then                       'auto scan => auto feedback gateway connect status (but step by step receive will not,so it wiil not use below method)
                If lSetBestPoll = False Then               'bother setting best polling(auto scan too fast)
                    If ret = 6 Or ret = 7 Then             'gateway response fastly
                        If ret = 6 Then
                            Form1.ListBox1.Items(nGw - 1) = "Gw" & nGw & ":" & GwConf(nGw - 1).ip & " Waitting"
                        ElseIf ret = 7 Then
                            InitClearTags(nGw)
                            Form1.ListBox1.Items(nGw - 1) = "Gw" & nGw & ":" & GwConf(nGw - 1).ip & " ok"
                        End If
                    Else
                        Form1.ListBox1.Items(nGw - 1) = "Gw" & nGw & ":" & GwConf(nGw - 1).ip & " " & AB_ErrMsg(ret)
                    End If
                End If
            Else
                Form1.ListBox1.Items(nGw - 1) = "Gw" & nGw & ":" & GwConf(nGw - 1).ip & " " & AB_ErrMsg(ret)
            End If
        Next nGw

        If lAutoScan = True Then
            BlinkAutoStatus()
        End If
        Timer1.Interval = 100 'ScanTimer
        Timer1.Enabled = True
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Try
            Dim node As Integer, ret As Integer
            node = CInt(252)
            Dim nLamp As Byte = Val(Mid(1, 1, 1))
            Dim nColor As Byte = Val(Mid(0, 1, 1))
            ret = AB_LED_Status(CurGwId, node, 0, nLamp)
        Catch ex As Exception
            WriteErrorx("[controller problem]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "-" & ex.Message)
            MsgBox("KONEKSI KE CONTROLLER GAGAL | CEK KABEL JARINGAN ANDA")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        write_pos1()
        Call panggil(1, TextBox4.Text)
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs)
        Call keadaan_on(2)
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            write_pos2()
            Call panggil(2, TextBox5.Text)
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            write_pos1()
            Call panggil(1, TextBox4.Text)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Call keadaan_off(2)
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs)
        Call keadaan_on(3)
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs)
        Call keadaan_off(3)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs)
        Call keadaan_on(4)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs)
        Call keadaan_off(4)
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs)
        Call keadaan_off(5)
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs)
        Call keadaan_on(5)
    End Sub
    Private Sub Button27_Click(sender As Object, e As EventArgs)
        Call keadaan_off(6)
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs)
        Call keadaan_on(6)
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs)
        Call keadaan_off(7)
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs)
        Call keadaan_on(7)
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs)
        Call keadaan_on(8)
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs)
        Call keadaan_off(8)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        write_pos2()
        Call panggil(2, TextBox5.Text)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        write_pos3()
        Call panggil(3, TextBox9.Text)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        write_pos4()
        Call panggil(4, TextBox13.Text)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        write_pos5()
        Call panggil(5, TextBox17.Text)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        write_pos6()
        Call panggil(6, TextBox21.Text)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        write_pos7()
        Call panggil(7, TextBox25.Text)
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        write_pos8()
        Call panggil(8, TextBox29.Text)
    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            write_pos3()
            Call panggil(3, TextBox9.Text)
        End If
    End Sub

    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            write_pos4()
            Call panggil(4, TextBox13.Text)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        writeini(information_log, "LOG", "", "ERROR")
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles log7.Click

    End Sub

    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            write_pos5()
            Call panggil(5, TextBox17.Text)
        End If
    End Sub







    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            write_pos6()
            Call panggil(6, TextBox21.Text)
        End If
    End Sub

    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            write_pos7()
            Call panggil(7, TextBox25.Text)
        End If
    End Sub

    Private Sub TextBox29_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox29.KeyDown
        If e.KeyCode = Keys.Enter Then
            write_pos8()
            Call panggil(8, TextBox29.Text)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        liffting_default()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub
End Class