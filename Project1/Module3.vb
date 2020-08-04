Option Strict Off
Option Explicit On
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Text
Imports System.Data.Odbc
Module Module3
    Public DiagGWTime() As Single
    Public lDiagGwOk() As Boolean
    Public lSetBestPoll As Boolean, lCheckTagLive As Boolean
    Public CurGwId As Integer, DelayItem As String, CurNode As Long
    Public GwPort, NodeAddress As Short
    Public TagCommand As Short
    Public RcvTagData, SubCommandMess As String
    Public max_node, GwID, maxid As Short, mnuTagDiag As Boolean
    Public GwConf() As GwConfType
    Public koneksi_status_andon, koneksi_status_scan As Boolean
    'value content
    Public numvalue(8) As Integer
    Public FilePath As String = Application.StartupPath & "\history.ini"
    Public information_log As String = Application.StartupPath & "\information-log.ini"
    Public strin As String
    Dim objIniFile As New clsIni(My.Application.Info.DirectoryPath & "\conf.txt")
    Public hasil As String
    Public cmd As OdbcCommand
    Public rd As OdbcDataReader
    Public rd2 As OdbcDataReader
    Public adapter As OdbcDataAdapter
    Public dataset As New DataSet
    Public connection As OdbcConnection
    Public connection_scan As OdbcConnection
    Public connection_andon As OdbcConnection
    Public server, uid, pwd, databases, port As String
    Public server_scan, uid_scan, pwd_scan, databases_scan, port_scan As String
    Public server_andon, uid_andon, pwd_andon, databases_andon, port_andon As String
    Public Sub WriteErrorx(ByVal Error_Messagex As String)
        Dim objStreamWrite As StreamWriter
        objStreamWrite = New StreamWriter("information-log.txt", True, Encoding.Unicode)
        objStreamWrite.WriteLine(Error_Messagex)
        objStreamWrite.Close()
    End Sub
    Public Sub conf()
        server = objIniFile.GetString("pick", "server", "")
        uid = objIniFile.GetString("pick", "uid", "")
        pwd = objIniFile.GetString("pick", "pwd", "")
        databases = objIniFile.GetString("pick", "databases", "")
        port = objIniFile.GetString("pick", "port", "")

        server_scan = objIniFile.GetString("scan", "server_scan", "")
        uid_scan = objIniFile.GetString("scan", "uid_scan", "")
        pwd_scan = objIniFile.GetString("scan", "pwd_scan", "")
        databases_scan = objIniFile.GetString("scan", "databases_scan", "")
        port_scan = objIniFile.GetString("scan", "port_scan", "")

        server_andon = objIniFile.GetString("andon", "server_andon", "")
        uid_andon = objIniFile.GetString("andon", "uid_andon", "")
        pwd_andon = objIniFile.GetString("andon", "pwd_andon", "")
        databases_andon = objIniFile.GetString("andon", "databases_andon", "")
        port_andon = objIniFile.GetString("andon", "port_andon", "")
    End Sub
    Public Sub koneksi_andon()
        Call conf()
        Try
            connection_andon = New OdbcConnection("Driver=MySQL ODBC 5.1 Driver;SERVER=" & server_andon & ";UID=" & uid_andon & ";pwd=" & pwd_andon & ";DATABASE=" & databases_andon & ";PORT=" & port_andon & "")
            'str = "dsn=pis_2;server=10.59.4.107;uid=farel;pwd=Cryptonesia22;database=pis_2;"
            If connection_andon.State = ConnectionState.Closed Then
                connection_andon.Open()
                koneksi_status_andon = True
            End If
        Catch ex As Exception
            koneksi_status_andon = False
            WriteErrorx("[JARINGAN PROBLEM DB SAVE HISTORY ANDON NOT CONNECTION]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "-" & ex.Message)
        End Try
    End Sub
    Public Sub koneksi_scan()
        Call conf()
        Try
            connection_scan = New OdbcConnection("Driver=MySQL ODBC 5.1 Driver;SERVER=" & server_scan & ";UID=" & uid_scan & ";pwd=" & pwd_scan & ";DATABASE=" & databases_scan & ";PORT=" & port_scan & "")
            'str = "dsn=pis_2;server=10.59.4.107;uid=farel;pwd=Cryptonesia22;database=pis_2;"
            If connection_scan.State = ConnectionState.Closed Then
                connection_scan.Open()
                koneksi_status_scan = True
            End If
        Catch ex As Exception
            koneksi_status_scan = False
            WriteErrorx("[JARINGAN PROBLEM SCAN TABLE NOT CONNECTION]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "-" & ex.Message)
        End Try
    End Sub
    Public Sub koneksi()
        Call conf()
        Try
            connection = New OdbcConnection("Driver=MySQL ODBC 5.1 Driver;SERVER=" & server & ";UID=" & uid & ";pwd=" & pwd & ";DATABASE=" & databases & ";PORT=" & port & "")
            'str = "dsn=picking;server=10.59.4.107;uid=farel;pwd=Cryptonesia22;database=picking_lamp_trial;"
            'membuka koneksi
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If
            koneksi_status_scan = True
        Catch ex As Exception
            koneksi_status_scan = False
            WriteErrorx("[JARINGAN PROBLEM DB PIS_2 NOT CONNECTION]" & "-" & Format(Now, "[yyyy-MM-dd HH:mm:ss]") & "-" & ex.Message)
        End Try
    End Sub
    'value content
    Sub kurangvalue(pos, a)
        numvalue(pos) = numvalue(pos) - a
    End Sub
    Sub tambahvalue(pos, a)
        numvalue(pos) = numvalue(pos) + a
    End Sub
    Sub value(pos)
        Dim ret As Integer
        'confirm
       
        If pos = "1" Then
            ret = AB_LB_DspAddr(CurGwId, 180)
            ret = AB_LED_Status(CurGwId, 180, 2, 1)
            ret = AB_LB_SetLock(CurGwId, 180, 1, 1) ' mengunci konfirmasi
        ElseIf pos = "2" Then
            ret = AB_LB_DspAddr(CurGwId, 182)
            ret = AB_LED_Status(CurGwId, 182, 2, 1)
            ret = AB_LB_SetLock(CurGwId, 182, 1, 1)
        ElseIf pos = "3" Then
            ret = AB_LB_DspAddr(CurGwId, 184)
            ret = AB_LED_Status(CurGwId, 184, 2, 1)
            ret = AB_LB_SetLock(CurGwId, 184, 1, 1)
        ElseIf pos = "4" Then
            ret = AB_LB_DspAddr(CurGwId, 186)
            ret = AB_LED_Status(CurGwId, 186, 2, 1)
            ret = AB_LB_SetLock(CurGwId, 186, 1, 1)
        ElseIf pos = "5" Then
            ret = AB_LB_DspAddr(CurGwId, 188)
            ret = AB_LED_Status(CurGwId, 188, 2, 1)
            ret = AB_LB_SetLock(CurGwId, 188, 1, 1)
        ElseIf pos = "6" Then
            ret = AB_LB_DspAddr(CurGwId, 190)
            ret = AB_LED_Status(CurGwId, 190, 2, 1)
            ret = AB_LB_SetLock(CurGwId, 190, 1, 1)
        ElseIf pos = "7" Then
            ret = AB_LB_DspAddr(CurGwId, 192)
            ret = AB_LED_Status(CurGwId, 192, 2, 1)
            ret = AB_LB_SetLock(CurGwId, 192, 1, 1)
        ElseIf pos = "8" Then
            ret = AB_LB_DspAddr(CurGwId, 194)
            ret = AB_LED_Status(CurGwId, 194, 2, 1)
            ret = AB_LB_SetLock(CurGwId, 194, 1, 1)
        End If

        'confirm
    End Sub

    Structure GwConfType
        Dim gw_id As Short
        Dim ip As String
        Dim port As Short
    End Structure

    Function Bin2Str(ByRef bufbin() As Byte, ByRef start As Short, ByRef cnt As Object) As String
        Dim i As Short
        Dim bufstr As String
        bufstr = ""
        For i = 0 To cnt - 1
            bufstr = bufstr & Chr(bufbin(start + i))
        Next i

        Bin2Str = bufstr
    End Function
    Function AB_LoadConf() As Integer
        Dim ret As Integer
        Dim i As Integer, j As Integer
        Dim id As Integer, ip(20) As Byte, port As Integer, tmpstr As String

        GwCnt = AB_GW_Cnt()
        If GwCnt = 0 Then Exit Function
        ReDim GwConf(GwCnt - 1)

        For i = 0 To GwCnt - 1
            ret = AB_GW_Conf(i, id, ip(0), port)
            If ret >= 0 Then
                GwConf(i).gw_id = id
                GwConf(i).port = port
                tmpstr = ""
                For j = 0 To 19
                    If ip(j) = 0 Then Exit For
                    tmpstr = tmpstr + Chr(ip(j))
                Next j
                GwConf(i).ip = tmpstr

            Else
                GwConf(i).gw_id = 0
                GwConf(i).port = 0
                GwConf(i).ip = ""
            End If
        Next i
        AB_LoadConf = GwCnt
    End Function

    Function BinCopy(ByRef dstbuf() As Byte, ByRef dst_ndx As Short, ByRef srcbuf() As Byte, ByRef src_ndx As Short, ByRef cnt As Short) As Short
        Dim i As Short

        For i = 0 To cnt - 1
            dstbuf(dst_ndx + i) = srcbuf(src_ndx + i)
        Next i

        BinCopy = 0
    End Function

    Function Str2Bin(ByRef strdata As String, ByRef bufbin() As Byte, ByRef start As Short) As Integer
        '// change string to byte array
        Dim i As Short
        Dim strcnt, ndx, data As Integer

        strcnt = Len(strdata)
        ndx = start
        For i = 1 To strcnt
            data = Asc(Mid(strdata, i, 1))
            bufbin(ndx) = (data Mod 256) And &HFFS
            ndx = ndx + 1
            If data >= 256 Then
                bufbin(ndx) = (data \ 256) And &HFFS
                ndx = ndx + 1
            End If
        Next i

        Str2Bin = ndx
    End Function

    Function QwayGetDataB() As Short
        'ret>=0:ok,else:err
        Dim nodesample As Integer
        Dim ret As Integer
        Dim nMsg_type As Short
        Dim ccbdata(256) As Byte
        Dim ccblen As Integer, nRcvbun As Integer
        Dim k As Short
        Dim tmpstr, tmpstr1 As String, RcvTagData As String
        Dim tmp As Byte
        Dim nGwId As Integer, nNode As Short, TagCommand As Short
        nGwId = 0          '0: get any gateway status
        GwPort = 0
        nNode = 0
        TagCommand = -1
        nMsg_type = -1
        SubCommandMess = ""
        tmpstr1 = ""
        ccblen = 0
        nRcvbun = 0

        ret = AB_Tag_RcvMsg(nGwId, nNode, TagCommand, nMsg_type, ccbdata(0), ccblen)
        If ret > 0 Then
            If nNode < 0 Then
                GwPort = 1
            ElseIf nNode > 0 Then
                GwPort = 2
            End If
            nNode = Math.Abs(nNode)

            RcvTagData = Trim(Bin2Str(ccbdata, 0, ccblen))      'By AB_Tag_RcvMsg Funs.
            Select Case TagCommand
                Case 100
                    Select Case nMsg_type
                        Case 0
                            nRcvbun = AB_GW_RcvButton(ccbdata(0), ccblen)
                            '1   && Only CONFIRM button
                            '2   && Only UP button
                            '3   && Only DOWN button
                            '4   && First UP, then CONFIRM
                            '5   && First DOWN , then CONFIRM
                            '6   && First UP and DOWN , then CONFIRM
                            '7   && First DOWN , then UP
                            '8   && First CONFIRM , then UP
                            '9   && First DOWN and CONFIRM , then UP
                            '10  && First UP , then DOWN
                            '11  && First CONFIRM , then DOWN
                            '12  && First UP and CONFIRM , then DOWN
                            '---------------------'

                            SubCommandMess = "keycode return/receive Button value:" & Str(nRcvbun) & ")"

                        Case 1
                            SubCommandMess = "tag Busy)"
                        Case 2
                            SubCommandMess = "first record disappear)"
                        Case 3
                            SubCommandMess = "last record confirm)"
                    End Select
                    If nMsg_type = -1 Then
                        SubCommandMess = "Null"
                    Else
                        SubCommandMess = Str(nMsg_type) & " (" & SubCommandMess
                    End If

                    RcvTagData = Trim(Bin2Str(ccbdata, 0, ccblen))
                    Form1.ListBox2.Items.Add(TimeString & " [" & Str(nGwId) & "," & Str(GwPort) & "," & Str(nNode) & "," & Str(TagCommand) & "," & SubCommandMess & "]," & Trim(RcvTagData))
                    Form1.ListBox2.SelectedIndex = Form1.ListBox2.Items.Count - 1
                    ' Form2.ListBox1.Items.Add(TimeString & " [" & Str(nGwId) & "," & Str(GwPort) & "," & Str(nNode) & "," & Str(TagCommand) & "," & SubCommandMess & "]," & Trim(RcvTagData))
                    'Form2.ListBox1.SelectedIndex = Form2.ListBox1.Items.Count - 1
                    Form2.ListBox1.Items.Add(TimeString & " [" & Str(nGwId) & "," & Str(GwPort) & "," & Str(nNode) & "," & Str(TagCommand) & "," & SubCommandMess & "]," & Trim(RcvTagData))
                    Form2.ListBox1.SelectedIndex = Form2.ListBox1.Items.Count - 1

                Case 9
                    lDiagGwOk(nGwId) = True 'check gateway connectting
                    tmpstr = ldump3(ccbdata, 0, ccblen)

                    tmpstr = ""
                    If mnuTagDiag = True Then
                        If CurGwId = nGwId Then
                            MsgBox2(" Tag Diagnosis as below ===>")
                        End If
                    End If
                    For k = 1 To 250
                        If (((k - 1) Mod 8) = 0) Then
                            tmp = ccbdata(3 + ((k - 1) \ 8))
                        End If
                        If (tmp Mod 2) = 1 Then
                            tmpstr = tmpstr & "0"
                        Else
                            tmpstr = tmpstr & "1"
                            maxid = k
                        End If
                        tmp = tmp \ 2
                    Next k

                    'easy read
                    If mnuTagDiag = True Then
                        If CurGwId = nGwId Then
                            For k = 1 To nNode
                                tmpstr1 = tmpstr1 & Mid(tmpstr, k, 1)
                                If k > 1 And (k Mod 10) = 0 Then
                                    tmpstr1 = tmpstr1 & ","
                                End If
                                If (k Mod 50) = 0 Then
                                    Form1.ListBox2.Items.Add(Space(18) & tmpstr1 & Chr(10))
                                    Form2.ListBox1.Items.Add(Space(18) & tmpstr1 & Chr(10))
                                    tmpstr1 = ""
                                End If
                            Next k
                            If Trim(tmpstr1) <> "" Then
                                Form1.ListBox2.Items.Add(Space(18) & tmpstr1 & Chr(10))
                            End If
                            Form1.ListBox2.Items.Add(Space(18) & " <<GW id:" & Str(nGwId) & ",Max Node = " & maxid & ",Current polling range = " & nNode & ">>")
                            Form1.ListBox2.SelectedIndex = Form1.ListBox2.Items.Count - 1

                            Form2.ListBox1.Items.Add(Space(18) & " <<GW id:" & Str(nGwId) & ",Max Node = " & maxid & ",Current polling range = " & nNode & ">>")
                            Form2.ListBox1.SelectedIndex = Form1.ListBox1.Items.Count - 1
                            mnuTagDiag = False
                        Else
                            If lDiagGwOk(CurGwId) = False Then 'connect fail
                                mnuTagDiag = False
                            End If
                        End If
                    End If

                    'set best polling range
                    If lSetBestPoll = True Then
                        If CurGwId = nGwId Then
                            ret = AB_GW_SetPollRang(nGwId, (-1) * maxid)
                            Form1.Timer2.Enabled = False
                            lSetBestPoll = False
                            Form1.ListBox2.Items.Add(TimeString & " Gateway = " & nGwId & " " & ",Max Node = " & maxid)
                            Form1.ListBox2.SelectedIndex = Form1.ListBox2.Items.Count - 1
                            Form2.ListBox1.Items.Add(TimeString & " Gateway = " & nGwId & " " & ",Max Node = " & maxid)
                            Form2.ListBox1.SelectedIndex = Form2.ListBox1.Items.Count - 1
                            Form1.Button7.Text = "Set best polling range"
                            Form1.Button7.Enabled = True
                        End If
                    End If

                Case 10
                    If lCheckTagLive = True Then
                        If CurGwId = nGwId Then
                            If (GwPort = 1) Then nNode = nNode * (-1)
                            MsgBox2("Tag has died or not existed [" & Str(nGwId) & "," & nNode & "]")
                            MsgBox("Tag has died or not existed [" & Str(nGwId) & "," & nNode & "]")
                            lCheckTagLive = False
                            Form1.Timer2.Enabled = False
                        End If
                    End If

                Case 60
                    RcvTagData = Trim(Bin2Str(ccbdata, 0, ccblen))
                    Form1.ListBox2.Items.Add(TimeString & " [" & Str(nGwId) & "," & Str(GwPort) & "," & Str(nNode) & "," & Str(TagCommand) & "," & SubCommandMess & "]," & Right(RcvTagData, 8))
                    Form1.ListBox2.SelectedIndex = Form1.ListBox2.Items.Count - 1

                Case Else
                    RcvTagData = Trim(Bin2Str(ccbdata, 1, ccblen)) 'menerima data
                    Dim nLamp As Byte = Val(Mid(1, 1, 1)) ' menentukan lampun nyala 
                    Dim nColor As Byte = Val(Mid(0, 1, 1)) ' membuat lampu jadi merah

                    Dim node_check As Integer
                    If Not menu_form.part_name(Str(nNode)) = "" Then
                        node_check = CInt(Left(menu_form.part_name(Str(nNode)), 1))
                        Call kurangvalue(node_check, 1)
                        If numvalue(node_check) <= 0 Then
                            hasil = "selesai"
                            numvalue(node_check) = 0
                            Form2.Label1.Text = numvalue(node_check)
                        Else
                            hasil = "belum"
                            Form2.Label1.Text = numvalue(node_check)
                        End If
                        If hasil = "selesai" Then
                            If Not menu_form.part_name(Str(nNode)) = "" Then
                                'confirm
                                Dim hasil As Integer = CInt(Left(menu_form.part_name(Str(nNode)), 1))
                                'CInt(Left(menu_form.part_name(Str(nNode))
                                If hasil = "1" Then
                                    nodesample = 180
                                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                                    ret = AB_LED_Status(CurGwId, nodesample, 5, 3)
                                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                                ElseIf hasil = "2" Then
                                    nodesample = 182
                                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                                    ret = AB_LED_Status(CurGwId, nodesample, 5, 3)
                                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                                ElseIf hasil = "3" Then
                                    nodesample = 184
                                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                                    ret = AB_LED_Status(CurGwId, nodesample, 5, 3)
                                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                                ElseIf hasil = "4" Then
                                    nodesample = 186
                                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                                    ret = AB_LED_Status(CurGwId, nodesample, 5, 3)
                                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                                ElseIf hasil = "5" Then
                                    nodesample = 188
                                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                                    ret = AB_LED_Status(CurGwId, nodesample, 5, 3)
                                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                                ElseIf hasil = "6" Then
                                    nodesample = 190
                                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                                    ret = AB_LED_Status(CurGwId, nodesample, 5, 3)
                                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                                ElseIf hasil = "7" Then
                                    nodesample = 192
                                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                                    ret = AB_LED_Status(CurGwId, nodesample, 5, 3)
                                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                                ElseIf hasil = "8" Then
                                    nodesample = 194
                                    ret = AB_LB_DspAddr(CurGwId, nodesample)
                                    ret = AB_LED_Status(CurGwId, nodesample, 5, 3)
                                    ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                                    'confirm
                                End If
                            End If
                        End If
                    End If
                    ret = AB_LED_Status(nGwId, nNode, nColor, nLamp)
                    'release node
                    Dim a As Integer
                    'monitoring.ListBox1.Items.Add("1 | " & Str(nNode) & " | " & menu_form.part_name(a) & " > " & Trim(RcvTagData))
                    ' monitoring.ListBox1.Items.Add(TimeString & " [" & Str(nGwId) & "," & Str(GwPort) & "," & Str(nNode) & "," & Str(TagCommand) & "," & SubCommandMess & "]," & Trim(RcvTagData))
                    If node_check = 1 Then
                        monitoring.ListBox1.Items.Remove(menu_form.part_name(Str(nNode)))
                        menu_form.part_name(Str(nNode)) = "" ' release node content
                        'monitoring.andon_count(Str(nNode), monitoring.jum_part(1, Str(nNode)))
                    ElseIf node_check = 2 Then
                        monitoring.ListBox2.Items.Remove(menu_form.part_name(Str(nNode)))
                        menu_form.part_name(Str(nNode)) = "" ' release node content
                        'monitoring.andon_count(Str(nNode), monitoring.jum_part(2, Str(nNode)))
                    ElseIf node_check = 3 Then
                        monitoring.ListBox3.Items.Remove(menu_form.part_name(Str(nNode)))
                        menu_form.part_name(Str(nNode)) = "" ' release node content
                        'monitoring.andon_count(Str(nNode), monitoring.jum_part(3, Str(nNode)))
                    ElseIf node_check = 4 Then
                        monitoring.ListBox4.Items.Remove(menu_form.part_name(Str(nNode)))
                        menu_form.part_name(Str(nNode)) = "" ' release node content
                        'monitoring.andon_count(Str(nNode), monitoring.jum_part(4, Str(nNode)))
                    ElseIf node_check = 5 Then
                        monitoring.ListBox5.Items.Remove(menu_form.part_name(Str(nNode)))
                        menu_form.part_name(Str(nNode)) = "" ' release node content
                        'monitoring.andon_count(Str(nNode), monitoring.jum_part(5, Str(nNode)))
                    ElseIf node_check = 6 Then
                        monitoring.ListBox6.Items.Remove(menu_form.part_name(Str(nNode)))
                        menu_form.part_name(Str(nNode)) = "" ' release node content
                        'monitoring.andon_count(Str(nNode), monitoring.jum_part(6, Str(nNode)))
                    ElseIf node_check = 7 Then
                        monitoring.ListBox7.Items.Remove(menu_form.part_name(Str(nNode)))
                        menu_form.part_name(Str(nNode)) = "" ' release node content
                        'monitoring.andon_count(Str(nNode), monitoring.jum_part(7, Str(nNode)))
                    ElseIf node_check = 8 Then
                        monitoring.ListBox8.Items.Remove(menu_form.part_name(Str(nNode)))
                        menu_form.part_name(Str(nNode)) = "" ' release node content
                        'monitoring.andon_count(Str(nNode), monitoring.jum_part(8, Str(nNode)))
                    End If
                    'next liffting confirm button

                    If nNode = 180 Then
                        Dim liffting1 As Integer = CInt(monitoring.TextBox4.Text)
                        If liffting1 < 9999 Then
                            liffting1 = liffting1 + 1
                        Else
                            liffting1 = 0
                            liffting1 = liffting1 + 1
                        End If
                        monitoring.save_history(1)
                        monitoring.TextBox4.Text = liffting1
                        Call monitoring.panggil(1, liffting1)
                        liffting1 = 0
                        monitoring.write_pos1()
                    ElseIf nNode = 182 Then
                        Dim liffting2 As Integer = CInt(monitoring.TextBox5.Text)
                        If liffting2 < 9999 Then
                            liffting2 = liffting2 + 1
                        Else
                            liffting2 = 0
                            liffting2 = liffting2 + 1
                        End If
                        monitoring.save_history(2)
                        monitoring.TextBox5.Text = liffting2
                        Call monitoring.panggil(2, liffting2)
                        liffting2 = 0
                        monitoring.write_pos2()
                    ElseIf nNode = 184 Then
                        Dim liffting3 As Integer = CInt(monitoring.TextBox9.Text)
                        If liffting3 < 9999 Then
                            liffting3 = liffting3 + 1
                        Else
                            liffting3 = 0
                            liffting3 = liffting3 + 1
                        End If
                        monitoring.save_history(3)
                        monitoring.TextBox9.Text = liffting3
                        Call monitoring.panggil(3, liffting3)
                        liffting3 = 0
                        monitoring.write_pos3()
                    ElseIf nNode = 186 Then
                        Dim liffting4 As Integer = CInt(monitoring.TextBox13.Text)
                        If liffting4 < 9999 Then
                            liffting4 = liffting4 + 1
                        Else
                            liffting4 = 0
                            liffting4 = liffting4 + 1
                        End If
                        monitoring.save_history(4)
                        monitoring.TextBox13.Text = liffting4
                        Call monitoring.panggil(4, liffting4)
                        liffting4 = 0
                        monitoring.write_pos4()
                    ElseIf nNode = 188 Then
                        Dim liffting5 As Integer = CInt(monitoring.TextBox17.Text)
                        If liffting5 < 9999 Then
                            liffting5 = liffting5 + 1
                        Else
                            liffting5 = 0
                            liffting5 = liffting5 + 1
                        End If
                        monitoring.save_history(5)
                        monitoring.TextBox17.Text = liffting5
                        Call monitoring.panggil(5, liffting5)
                        liffting5 = 0
                        monitoring.write_pos5()
                    ElseIf nNode = 190 Then
                        Dim liffting6 As Integer = CInt(monitoring.TextBox21.Text)
                        If liffting6 < 9999 Then
                            liffting6 = liffting6 + 1
                        Else
                            liffting6 = 0
                            liffting6 = liffting6 + 1
                        End If
                        monitoring.save_history(6)
                        monitoring.TextBox21.Text = liffting6
                        Call monitoring.panggil(6, liffting6)
                        liffting6 = 0
                        monitoring.write_pos6()
                    ElseIf nNode = 192 Then
                        Dim liffting7 As Integer = CInt(monitoring.TextBox25.Text)
                        If liffting7 < 9999 Then
                            liffting7 = liffting7 + 1
                        Else
                            liffting7 = 0
                            liffting7 = liffting7 + 1
                        End If
                        monitoring.save_history(7)
                        monitoring.TextBox25.Text = liffting7
                        Call monitoring.panggil(7, liffting7)
                        liffting7 = 0
                        monitoring.write_pos7()
                    ElseIf nNode = 194 Then
                        Dim liffting8 As Integer = CInt(monitoring.TextBox29.Text)
                        If liffting8 < 9999 Then
                            liffting8 = liffting8 + 1
                        Else
                            liffting8 = 0
                            liffting8 = liffting8 + 1
                        End If
                        monitoring.save_history(8)
                        monitoring.TextBox29.Text = liffting8
                        Call monitoring.panggil(8, liffting8)
                        liffting8 = 0
                        monitoring.write_pos8()
                    End If
                    'ListBox1.Items.Add(menu_form.part_name(a))
                    'a = a + 1
                    'contoh
                    'Form2.ListBox1.Items.Add(TimeString & " [" & Str(nGwId) & "," & Str(GwPort) & "," & Str(nNode) & "," & Str(TagCommand) & "," & SubCommandMess & "]," & Trim(RcvTagData))
                    'Call monitoring.panggil()
                    'Form1.ListBox2.Items.Add(TimeString & " [" & Str(nGwId) & "," & Str(GwPort) & "," & Str(nNode) & "," & Str(TagCommand) & "," & SubCommandMess & "]," & Trim(RcvTagData))
                    'Form1.ListBox2.SelectedIndex = Form1.ListBox2.Items.Count - 1

            End Select
        End If
        QwayGetDataB = ret
    End Function
    Function DisplayNum(ByVal val As Long) As Boolean
        Dim nGw As Integer, ret As Integer
        For nGw = 1 To GwCnt
            ret = AB_LB_DspNum(nGw, -252, val, 0, 0)
        Next nGw
    End Function
    Public Sub clearTags()
        Dim nGw As Integer, ret As Integer
        For nGw = 1 To GwCnt
            ret = AB_LB_DspStr(nGw, -252, 0, 0, -2)
            ret = AB_LB_DspStr(nGw, 252, 0, 0, -2)
            AB_BUZ_On(nGw, 252, 0)
            AB_BUZ_On(nGw, -252, 0)
            AB_LED_Status(nGw, 252, 0, 0)
            AB_LED_Status(nGw, -252, 0, 0)

            ret = AB_AHA_ClrDsp(nGw, 252)
            ret = AB_AHA_ClrDsp(nGw, -252)
            AB_AHA_BUZ_On(nGw, 252, 0)
            AB_AHA_BUZ_On(nGw, -252, 0)
            AB_AHA_LED_Dsp(nGw, 252, 0, 0)
            AB_AHA_LED_Dsp(nGw, -252, 0, 0)

            ret = AB_DCS_Cls(nGw, 252)
            ret = AB_DCS_Cls(nGw, -252)
        Next
    End Sub
    Public Sub TimeDelay(ByVal DT As Integer)
        Dim StartTick As Integer
        StartTick = Environment.TickCount() '在這邊記下一個數字
        Do
            If Environment.TickCount() - StartTick >= DT Then Exit Do '這段是判斷是否經過了DT個微秒數, 如果是就離開, DT是你給的傳入值
            'Application.DoEvents() '如果還沒到你希望的延遲時間, 先讓系統去處理其他等待中的工作
        Loop
    End Sub

End Module