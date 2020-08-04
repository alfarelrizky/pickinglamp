Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.Odbc
Imports System.Runtime.InteropServices
Public Class testlamp
    Dim cmd As OdbcCommand
    Dim rd As OdbcDataReader
    Dim rd2 As OdbcDataReader
    Dim adapter As OdbcDataAdapter
    Dim dataset As New DataSet
    Dim connection As OdbcConnection
    Dim connection_scan As OdbcConnection
    Dim str As String
    Dim part_name As String
    Dim server, uid, pwd, databases, port As String
    Dim server_scan, uid_scan, pwd_scan, databases_scan, port_scan As String
    Dim objIniFile As New clsIni(My.Application.Info.DirectoryPath & "\conf.txt")

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
    End Sub
    Sub koneksi()
        Call conf()
        connection = New OdbcConnection("Driver=MySQL ODBC 5.1 Driver;SERVER=" & server & ";UID=" & uid & ";pwd=" & pwd & ";DATABASE=" & databases & ";PORT=" & port & "")
        'str = "dsn=picking;server=10.59.4.107;uid=farel;pwd=Cryptonesia22;database=picking_lamp_trial;"
        'membuka koneksi
        If connection.State = ConnectionState.Closed Then
            Try
                connection.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox("database bermasalah", MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Sub sub_panggil(pos, suffix, model_code, color_code)
        Dim nod As Integer = 0
        Dim a As Integer
        Dim node As Integer
        Dim ret As Integer, nLamp As Byte
        Module3.numvalue(pos) = 0
        str = "select *from tb_master_pick_sps_final where pos = '" & ComboBox1.Text & "' and suffix = '" & TextBox3.Text & "' and type = '" & TextBox1.Text & "' and color_code = '" & TextBox2.Text & "' and jum_part <> '0'"
        cmd = New OdbcCommand(str, connection)
        rd = cmd.ExecuteReader()
        While rd.Read()
            nLamp = Val(Mid(1, 1, 2))
            ret = AB_LB_DspAddr(CurGwId, rd("node"))
            ret = AB_LED_Status(CurGwId, rd("node"), 1, nLamp)
            ret = AB_LB_DspNum(CurGwId, rd("node"), rd("Jum_part"), 2, 0)
            node = rd("node")
            part_name = pos & " | " & rd("node") & " | " & rd("part_name")
            ListBox1.Items.Add(part_name)
        End While
        If Module3.numvalue(pos) = 0 Then
            'MsgBox("ga ada")
            Dim nodesample As Integer
            If pos = "1" Then
                nodesample = 180
                ret = AB_LB_DspAddr(CurGwId, nodesample)
                ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
            ElseIf pos = "2" Then
                nodesample = 182
                ret = AB_LB_DspAddr(CurGwId, nodesample)
                ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
            ElseIf pos = "3" Then
                nodesample = 184
                ret = AB_LB_DspAddr(CurGwId, nodesample)
                ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
            ElseIf pos = "4" Then
                nodesample = 186
                ret = AB_LB_DspAddr(CurGwId, nodesample)
                ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
            ElseIf pos = "5" Then
                nodesample = 188
                ret = AB_LB_DspAddr(CurGwId, nodesample)
                ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
            ElseIf pos = "6" Then
                nodesample = 190
                ret = AB_LB_DspAddr(CurGwId, nodesample)
                ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
            ElseIf pos = "7" Then
                nodesample = 192
                ret = AB_LB_DspAddr(CurGwId, nodesample)
                ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
            ElseIf pos = "8" Then
                nodesample = 194
                ret = AB_LB_DspAddr(CurGwId, nodesample)
                ret = AB_LED_Status(CurGwId, nodesample, 1, 3)
                ret = AB_LB_SetLock(CurGwId, nodesample, 0, 1)
                'confirm
            End If
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub
    Private Sub testlamp_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call koneksi()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Clear()
        Dim node As Integer, ret As Integer
        node = CInt(252)
        ret = AB_LB_DspNum(CurGwId, node, 1, 1, -3)
        ret = AB_AHA_ClrDsp(CurGwId, node)
        Call monitoring.bersih()
        Call sub_panggil(ComboBox1.Text, TextBox3.Text, TextBox1.Text, TextBox2.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim node As Integer, ret As Integer
        node = CInt(252)
        ret = AB_LB_DspNum(CurGwId, node, 1, 1, -3)
        ret = AB_AHA_ClrDsp(CurGwId, node)
        Call monitoring.bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        ListBox1.Items.Clear()
        Dim node As Integer
        node = CInt(252)
        AB_TAG_Reset(CurGwId, node)
        
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim node As Integer, ret As Integer
        node = CInt(252)
        ret = AB_LB_DspNum(CurGwId, node, 1, 1, -3)
        ret = AB_AHA_ClrDsp(CurGwId, node)
        Call monitoring.bersih()
        TextBox4.Text = ""
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim ret As Integer
        Dim nLamp As Byte
        Dim a As Integer
        a = CInt(TextBox4.Text)
        nLamp = Val(Mid(1, 1, 2))
        ret = AB_LB_DspAddr(CurGwId, a)
        ret = AB_LED_Status(CurGwId, a, 1, nLamp)
    End Sub
End Class