Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.Odbc
Public Class Form2
    Inherits System.Windows.Forms.Form
    'database
    Dim cmd As OdbcCommand
    Dim rd As OdbcDataReader
    Dim rd2 As OdbcDataReader
    Dim adapter As OdbcDataAdapter
    Dim dataset As New DataSet
    Dim connection As OdbcConnection
    Dim str As String
    'database

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
    Sub koneksi()
        str = "dsn=picking;server=localhost;uid=root;pwd=;database=renah;"
        'membuka koneksi
        connection = New OdbcConnection(str)
        If connection.State = ConnectionState.Closed Then
            Try
                connection.Open()
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox("database bermasalah", MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Sub panggil()
        Dim ret As Integer, nLamp As Byte
        'Module3.numvalue = 0
        str = "select *from picking_lamp where simbol = 'a'"
        cmd = New OdbcCommand(str, connection)
        rd = cmd.ExecuteReader()
        While rd.Read()
            'Call Module3.tambahvalue(1)
            nLamp = Val(Mid(1, 1, 2))
            ret = AB_LB_DspAddr(CurGwId, rd("node"))
            ret = AB_LED_Status(CurGwId, rd("node"), 5, nLamp)
            Call Module3.value(rd("node"))
        End While
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
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call koneksi()
        Timer1.Enabled = True
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

        'waktu itung mundur
        Timer2.Enabled = True
        'waktu itung mundur
    End Sub
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call panggil()
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

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim ret As Short
        Timer2.Tag = CDbl(Timer2.Tag) - 1
        If CDbl(Timer2.Tag) <= 0 Then
            Timer2.Enabled = False
        End If
        Select Case DelayItem
            Case "reconnect"
                'reconnect delay
                ret = AB_GW_Open(CurGwId)
                Timer2.Enabled = False

            Case "SetBestPoll"
                lSetBestPoll = True
                ret = AB_GW_TagDiag(CurGwId, 1)
                Timer2.Enabled = False

            Case "CheckTagLive"
                If CDbl(Timer2.Tag) = 0 Then
                    If lDiagGwOk(CurGwId) = True Then 'it must not connect fail
                        ret = AB_LED_Status(CurGwId, CurNode, 0, 0)
                        MsgBox2("Tag live [" & Str(CurGwId) & "," & CurNode & "]")
                        MsgBox("Tag live [" & Str(CurGwId) & "," & CurNode & "]")
                        lCheckTagLive = False
                    End If
                End If
            Case "CountTing"
                DisplayNum(111111 * dispnum)
                dispnum = dispnum + 1
                If (dispnum > 9) Then dispnum = 1
                Timer2.Enabled = True
        End Select

        'itung mundur
        'Label1.Text = numvalue
        Timer2.Interval = 100
        'Timer2.Enabled = True
        'waktu = waktu - 1
        ' Label1.Text = waktu
        ' If waktu = 0 Then
        'waktu = 11
        'End If
        'itung mundur
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim node As Integer, ret As Integer
        node = CInt(252)
        ret = AB_LB_DspNum(CurGwId, node, 1, 1, -3)
        ret = AB_AHA_ClrDsp(CurGwId, node)

    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)
        Dim a As Integer = 200
        Dim b As Integer = 1
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs)
        Dim a As Integer = 200
        Dim b As Integer = 25
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
    End Sub
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ret As Short, nGw As Integer
        On Error Resume Next
        lAutoScan = False
        Timer1.Enabled = False
        For nGw = 1 To GwCnt
            ret = AB_GW_Status(nGw)
            If ret > 0 Then
                ret = AB_LB_DspStr(nGw, -252, CStr(0), 0, -3) 'turn off all tags on port1
                ret = AB_LB_DspStr(nGw, 252, CStr(0), 0, -3) 'turn off all tags on port2

                ret = AB_GW_Close(nGw) 'Close the specified Gateway
                ret = AB_API_Close()      'Shutdown the ABLEPICK API
            End If
        Next nGw
        End
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim ret As Integer, nLamp As Byte
        ' Module3.numvalue = 0
        str = "select *from picking_lamp where simbol = 'b'"
        cmd = New OdbcCommand(str, connection)
        rd = cmd.ExecuteReader()
        While rd.Read()
            ' Call Module3.tambahvalue(1)
            nLamp = Val(Mid(1, 1, 2))
            ret = AB_LB_DspAddr(CurGwId, rd("node"))
            ret = AB_LED_Status(CurGwId, rd("node"), 5, nLamp)
            ' ret = AB_LED_Dsp(CurGwId, rd("node"), nLamp, nLed)
            'Call Module3.value(())
        End While
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim ret As Integer, nLamp As Byte
        'Module3.numvalue = 0
        str = "select *from picking_lamp where simbol = 'c'"
        cmd = New OdbcCommand(str, connection)
        rd = cmd.ExecuteReader()
        While rd.Read()
            'Call Module3.tambahvalue(1)
            nLamp = Val(Mid(1, 1, 2))
            ret = AB_LB_DspAddr(CurGwId, rd("node"))
            ret = AB_LED_Status(CurGwId, rd("node"), 5, nLamp)
            ' ret = AB_LED_Dsp(CurGwId, rd("node"), nLamp, nLed)
            'Call Module3.value()
        End While
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim node As Integer, ret As Integer
        node = CInt(252)
        Dim nLamp As Byte = Val(Mid(1, 1, 1))
        Dim nColor As Byte = Val(Mid(0, 1, 1))
        ret = AB_LED_Status(CurGwId, node, 0, nLamp)
    End Sub
End Class