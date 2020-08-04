Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Public Class Form1
    Inherits System.Windows.Forms.Form
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

#Region "Windows Form 設計工具產生的程式碼 "
    Public Sub New()
        MyBase.New()
        If m_vb6FormDefInstance Is Nothing Then
            If m_InitializingDefInstance Then
                m_vb6FormDefInstance = Me
            Else
                Try
                    '就啟動表單而言，建立的第一個執行個體就是預設的執行個體。
                    If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
                        m_vb6FormDefInstance = Me
                    End If
                Catch
                End Try
            End If
        End If
        '此呼叫為 Windows Form 設計工具的必要項。
        InitializeComponent()
    End Sub
    'Form 覆寫 Dispose 以清除元件清單。
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
#End Region

#Region "升級支援 "
    Private Shared m_vb6FormDefInstance As Form1
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As Form1
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New Form1()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As Form1)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region

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

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ret As Short, nGw As Integer
        ChDrive(Mid(My.Application.Info.DirectoryPath, 1, 1))
        ChDir(My.Application.Info.DirectoryPath)
        ret = AB_API_Open() 'Startup and initialize the ABLELINK API
        LoadInitial()
        For nGw = 1 To GwCnt
            ListBox1.Items.Add("Gw" & nGw & ":" & GwConf(nGw - 1).ip)
        Next nGw

        If ret < 0 Then
            ListBox2.Items.Add(TimeString & " API Open Fail")
            ListBox2.SelectedIndex = ListBox2.Items.Count - 1
            Exit Sub
        Else
            ListBox2.Items.Add(TimeString & " API Open Successful")
            ListBox2.SelectedIndex = ListBox2.Items.Count - 1
        End If
        DelayItem = ""
        lAutoScan = True
        lSetBestPoll = False
        lInitClearTags = False
        dispnum = 1
        ComboBox1.Items.Add("-252")
        ComboBox1.Items.Add("+252")

        ComboBox2.Items.Add("123")
        ComboBox2.Items.Add("12345")
        ComboBox2.Items.Add("88888")

        ComboBox3.Items.Add(">0:blinking(by msec)")
        ComboBox3.Items.Add(" 0:Normal display")
        ComboBox3.Items.Add("-1:default blinking")
        ComboBox3.Items.Add("-2:Digits off")
        ComboBox3.Items.Add("-3:Digits and lamp off")

        ComboBox4.Items.Add("12345")
        ComboBox4.Items.Add("ABCDEF")
        ComboBox4.Items.Add("123456789012")
        ComboBox4.Items.Add("ABCDEFGHIJKL")
        ComboBox4.Items.Add("MNOPQRSTUVWX")
        ComboBox4.Items.Add("YZ!@#$%^&*()")
        ComboBox4.Items.Add("_-+={}[]|\:;")

        ComboBox5.Items.Add("-250")
        ComboBox5.Items.Add("250")

        ComboBox8.Items.Add("0:LED off")
        ComboBox8.Items.Add("1:LED on")
        ComboBox8.Items.Add("2:LED blinking")

        ComboBox7.Items.Add("0:Red")
        ComboBox7.Items.Add("1:Green")
        ComboBox7.Items.Add("2:Orange")

        ComboBox6.Items.Add("1:Buzzer on")
        ComboBox6.Items.Add("0:Buzzer off")

        ComboBox10.Items.Add("1:Jingle bells")
        ComboBox10.Items.Add("2:Carmen")
        ComboBox10.Items.Add("3:Happy Chinese new year")
        ComboBox10.Items.Add("4:Edelweiss")
        ComboBox10.Items.Add("5:Going home")
        ComboBox10.Items.Add("6:PAPALA")
        ComboBox10.Items.Add("7:Classical")
        ComboBox10.Items.Add("8:Listen to the rhythm of the falling rain")
        ComboBox10.Items.Add("9:Rock and roll")
        ComboBox10.Items.Add("10:Happy birthday")
        ComboBox10.Items.Add("11:Do Re Me")
        ComboBox10.Items.Add("12:Strauss")

        ComboBox11.Items.Add("0:0 seconds")
        ComboBox11.Items.Add("1:0.8 seconds")
        ComboBox11.Items.Add("2:1.6 seconds")
        ComboBox11.Items.Add("3:2.4 seconds")
        ComboBox11.Items.Add("4:3.2 seconds")
        ComboBox11.Items.Add("5:4.0 seconds")

        ComboBox12.Items.Add("0:Disabled")
        ComboBox12.Items.Add("1:Enabled")

        ComboBox13.Items.Add("1:Confirmation")
        ComboBox13.Items.Add("2:Shortage")

        ComboBox14.Items.Add("1:Confirm key")
        ComboBox14.Items.Add("2:Shortage key")
        ComboBox14.Items.Add("3:All key")

        ComboBox15.Items.Add("0:Unlock key")
        ComboBox15.Items.Add("1:Lock key")

        ComboBox16.Items.Add("5 digit")
        ComboBox16.Items.Add("4 digit")
        ComboBox16.Items.Add("3 digit")
        ComboBox16.Items.Add("2 digit")
        ComboBox16.Items.Add("1 digit")

        ComboBox17.Items.Add("0:Picking mode")
        ComboBox17.Items.Add("1:Stock mode")

        ComboBox18.Items.Add("0:All arrows off")
        ComboBox18.Items.Add("1:Right/up arrow on")
        ComboBox18.Items.Add("2:Left/down arrow on")
        ComboBox18.Items.Add("3:Both arrows on")

        ComboBox20.Items.Add("0:do not save EEPROM")
        ComboBox20.Items.Add("1:save into EEPROM")

        ComboBox21.Items.Add("1:Enable confirmed")
        ComboBox21.Items.Add("0:Disable confirmed")

        ComboBox22.Items.Add("1:Display on")
        ComboBox22.Items.Add("2:Blinking(2 sec)")
        ComboBox22.Items.Add("3:Blinking(1 sec)")
        ComboBox22.Items.Add("4:Blinking(0.5 sec)")
        ComboBox22.Items.Add("5:Blinking(0.25 sec)")

        ComboBox23.Items.Add("0:LED off")
        ComboBox23.Items.Add("1:LED on")
        ComboBox23.Items.Add("2:Blinking(2 sec)")
        ComboBox23.Items.Add("3:Blinking(1 sec)")
        ComboBox23.Items.Add("4:Blinking(0.5 sec)")
        ComboBox23.Items.Add("5:Blinking(0.25 sec)")

        ComboBox24.Items.Add("0:LED 1 work")
        ComboBox24.Items.Add("1:LED 2 work(reserved)")
        ComboBox24.Items.Add("2:Two LEDs work")

        ComboBox25.Items.Add("0:Buzzer off")
        ComboBox25.Items.Add("1:Buzzer on")
        ComboBox25.Items.Add("2:Freqency(slow)")
        ComboBox25.Items.Add("3:Freqency(slow)")
        ComboBox25.Items.Add("4:Freqency(quick)")
        ComboBox25.Items.Add("5:Freqency(quick)")

        ScanTimer = 100         'generally value = 100
        Timer1.Interval = ScanTimer
        CurGwId = 0   'Open All Gateway (Below in Dap_Setup())
        Dap_Setup()

        'Default 1st Gateway 
        ListBox1.SelectedIndex = 0
        CurGwId = 1
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim node As Integer, ret As Integer, value As Integer, nDot, nLamp, nColor As Byte
        Dim nLed As Integer
        RadioBtn2.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        value = CInt(Trim(ComboBox2.Text))
        nDot = DotConvert(CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, CheckBox7.Checked, CheckBox8.Checked)
        nLed = GetLEDStatus(Trim(ComboBox3.Text))
        nLamp = Val(Mid(ComboBox8.Text, 1, 1))
        nColor = Val(Mid(ComboBox7.Text, 1, 1))
        ret = AB_LB_DspNum(CurGwId, node, value, nDot, nLed)
    End Sub

    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
        Dim ret As Short, nGw As Integer
        On Error Resume Next

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
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim node As Integer, ret As Integer
        Me.RadioBtn1.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        ret = AB_LB_DspNum(CurGwId, node, 0, 0, -3)
        ret = AB_AHA_ClrDsp(CurGwId, node)        'clear AHA Display
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
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
                            ListBox1.Items(nGw - 1) = "Gw" & nGw & ":" & GwConf(nGw - 1).ip & " Waitting"
                        ElseIf ret = 7 Then
                            InitClearTags(nGw)
                            ListBox1.Items(nGw - 1) = "Gw" & nGw & ":" & GwConf(nGw - 1).ip & " ok"
                        End If
                    Else
                        ListBox1.Items(nGw - 1) = "Gw" & nGw & ":" & GwConf(nGw - 1).ip & " " & AB_ErrMsg(ret)
                    End If
                End If
            Else
                ListBox1.Items(nGw - 1) = "Gw" & nGw & ":" & GwConf(nGw - 1).ip & " " & AB_ErrMsg(ret)
            End If
        Next nGw

        If lAutoScan = True Then
            BlinkAutoStatus()
            Button9_Click(Button9, New System.EventArgs())
        End If
        Timer1.Interval = 100 'ScanTimer
        Timer1.Enabled = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim node As Integer, ret As Integer, value As String, nDot As Byte, nLed As Integer
        Me.RadioBtn3.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        value = Trim(ComboBox4.Text)
        nLed = GetLEDStatus(Trim(ComboBox3.Text))
        nDot = DotConvert(CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, CheckBox7.Checked, CheckBox8.Checked)
        ret = AB_LB_DspStr(CurGwId, node, value, nDot, nLed)
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim cnt As Integer, nGw As Integer
        'get received tag messages
        If lAutoScan = True Then
            For nGw = 1 To GwCnt
                If AB_GW_Status(nGw) = 7 Then
                    cnt = QwayGetDataB()           'API COMMAND
                End If
            Next nGw
        Else
            cnt = QwayGetDataB()
            If cnt = 0 Then
                ListBox2.Items.Add(TimeString & " Gateway = " & CurGwId & " " & ",No Receive Datas")
                ListBox2.SelectedIndex = ListBox2.Items.Count - 1
            End If
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
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

    End Sub

    Private Sub OpenBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenBtn.Click
        Dap_Setup()
    End Sub

    Private Sub CloseBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseBtn.Click
        Dim ret As Short

        ret = AB_LB_DspStr(CurGwId, -252, CStr(0), 0, -3) 'turn off all tags
        If AB_GW_Status(CurGwId) > 0 Then
            ret = AB_GW_Close(CurGwId) 'Close the specified Gateway
        End If
        ListBox2.Items.Add(TimeString & " Gateway = " & GwNode & " " & AB_ErrMsg(ret))
        ListBox2.SelectedIndex = ListBox2.Items.Count - 1
        'tmpfile = Shell(My.Application.Info.DirectoryPath & "\" & "darp.bat") 'clear ARP
        If ret < 0 Then
            Exit Sub
        End If

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedItem <> "" Then
            CurGwId = GwConf(ListBox1.SelectedIndex).gw_id
        End If
    End Sub

    Private Sub cmdClrList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClrList.Click
        ListBox2.Items.Clear()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Button4.Text = "Counting test" Then
            Button4.Text = "Off"
            DelayItem = "CountTing"
            Timer2.Interval = 500
            Timer2.Enabled = True
        Else
            Button4.Text = "Counting test"
            DelayItem = ""
            dispnum = 1
            Timer2.Enabled = False
            clearTags()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim ret As Integer
        ret = AB_GW_TagDiag(CurGwId, 1)
        mnuTagDiag = True
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Dim ret As Integer, nMaxNode As Integer
        RadioBtn9.Checked = True
        nMaxNode = Val(ComboBox5.Text)    'Sample:Port1(-),Port(+) //nMaxNode = Math.Abs(Val(ComboBox5.Text))
        If nMaxNode = 0 Then
            MsgBox("Please input the Max Node value !!")
        Else
            ret = AB_GW_SetPollRang(CurGwId, nMaxNode)
            ListBox2.Items.Add(TimeString & " Gateway = " & CurGwId & " " & ",Max Node = " + Trim(ComboBox5.Text))
            ListBox2.SelectedIndex = ListBox2.Items.Count - 1
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim node As Integer, ret As Integer
        Me.RadioBtn4.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        ret = AB_LB_DspAddr(CurGwId, node)
    End Sub

    Private Sub RadioBtn2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn2.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label2.ForeColor = Color.Red
        Label4.ForeColor = Color.Red
        GroupBox6.ForeColor = Color.Red
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim ret As Integer, nLamp As Byte, node As Integer, nLed As Integer
        Me.RadioBtn5.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nLamp = Val(Mid(ComboBox8.Text, 1, 1))
        nLed = GetLEDStatus(Trim(ComboBox3.Text))
        ret = AB_LED_Status(CurGwId, node, 5, nLamp)
        ret = AB_LED_Dsp(CurGwId, node, nLamp, nLed)
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim ret As Integer, nLamp As Integer, node As Integer, nColor As Integer
        Me.RadioBtn6.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nLamp = Val(Mid(ComboBox8.Text, 1, 1))
        nColor = Val(Mid(ComboBox7.Text, 1, 1))
        ret = AB_LED_Status(CurGwId, node, nColor, nLamp)
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim node As Integer
        Dim nOn As Integer
        RadioBtn7.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nOn = Val(Mid(ComboBox6.Text, 1, 1))
        AB_BUZ_On(CurGwId, node, nOn)
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim node As Integer
        RadioBtn8.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        AB_TAG_Reset(CurGwId, node)
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim ret As Integer, node As Integer, nMode As Integer
        RadioBtn10.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nMode = Val(Mid(ComboBox17.Text, 1, 1))
        ret = AB_LB_SetMode(CurGwId, node, nMode)
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        Dim ret As Integer, node As Integer, nDigit As Integer
        RadioBtn11.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nDigit = Val(Mid(ComboBox16.Text, 1, 1))
        ret = AB_TAG_CountDigit(CurGwId, node, nDigit)
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim ret As Integer, node As Integer, state As Integer, key As Integer
        RadioBtn12.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        state = Val(Mid(ComboBox15.Text, 1, 1))
        key = Val(Mid(ComboBox14.Text, 1, 1))
        ret = AB_LB_SetLock(CurGwId, node, state, key)    'lock shortage button condition: tag mode must disabled [shortage button] (ex:115,123)
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        Dim ret As Integer, node As Integer, nMode As Byte
        RadioBtn13.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nMode = Val(Mid(ComboBox13.Text, 1, 1))
        ret = AB_LB_Simulate(CurGwId, node, nMode)
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Dim ret As Integer, node As Integer, nTagMode As Byte, savemode As Integer
        RadioBtn14.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nTagMode = CaluTagModeValue()
        If Mid(ComboBox20.Text, 1, 1) = 0 Then
            savemode = 0
        Else
            savemode = 1
        End If
        If nTagMode > 0 Then
            ret = AB_TAG_mode(CurGwId, node, savemode, nTagMode)
            ListBox2.Items.Add(TimeString & " Gateway = " & CurGwId & " " & ",Node:" + Trim(Str(node)) & ",Save mode:" + Trim(Str(savemode)) & ",Tag mode:" + Trim(Str(nTagMode)))
            ListBox2.SelectedIndex = ListBox2.Items.Count - 1
        End If
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        Dim ret As Integer, node As Integer, nState As Integer
        RadioBtn15.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nState = Val(Mid(ComboBox12.Text, 1, 1))
        ret = AB_TAG_Complete(CurGwId, node, nState)
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim ret As Integer, node As Integer, nDelayItem As Integer, nDelayTime As Single
        RadioBtn16.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nDelayItem = Val(Mid(ComboBox11.Text, 1, 1))
        ret = AB_TAG_ButtonDelay(CurGwId, node, nDelayTime)
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Dim ret As Integer, node As Integer, cString As String, confirm As Byte, interval As Byte
        RadioBtn17.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        cString = Trim(ComboBox4.Text)
        confirm = Val(Mid(ComboBox21.Text, 1, 1))
        interval = Val(Mid(ComboBox22.Text, 1, 1))
        ret = AB_AHA_DspStr(CurGwId, node, Trim(cString), confirm, interval)
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Dim ret As Integer, node As Integer
        RadioBtn18.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        ret = AB_AHA_ReDsp(CurGwId, node)
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Dim ret As Integer, node As Integer, nLampType As Byte, nLampState As Integer
        RadioBtn19.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nLampType = Val(Mid(ComboBox24.Text, 1, 1))
        nLampState = Val(Mid(ComboBox23.Text, 1, 1))
        ret = AB_AHA_LED_Dsp(CurGwId, node, nLampType, nLampState)
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Dim ret As Integer, node As Integer, interval As Integer
        RadioBtn20.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        interval = Val(Mid(ComboBox25.Text, 1, 1))
        ret = AB_AHA_BUZ_On(CurGwId, node, interval)
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Dim ret As Integer, node As Integer, nSong As Byte, nBuzType As Byte
        RadioBtn21.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nSong = Val(Mid(ComboBox10.Text, 1, 1))
        nBuzType = Val(Mid(ComboBox6.Text, 1, 1))
        ret = AB_Melody_On(CurGwId, node, nSong, nBuzType)
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Dim ret As Integer, node As Integer, nVolumn As Byte
        RadioBtn22.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nVolumn = CInt(Trim(TextBox1.Text))
        ret = AB_Melody_Volume(CurGwId, node, nVolumn)
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Dim ret As Integer, node As Integer
        Dim value As Long, nArrow As Byte, nDot As Byte, interval As Integer
        RadioBtn23.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        value = CInt(Trim(ComboBox2.Text))
        nArrow = Val(Mid(ComboBox18.Text, 1, 1))
        nDot = DotConvert(CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, CheckBox7.Checked, CheckBox8.Checked)
        interval = GetLEDStatus(Trim(ComboBox3.Text))
        ret = AB_AV_LB_DspNum(CurGwId, node, value, nArrow, nDot, interval)
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Dim ret As Integer, node As Integer
        Dim cRow As String, cCol As String, value As Long, nDot As Integer, interval As Integer
        If Val(TextBox2.Text) >= 10 Then TextBox2.Text = "9"
        If Val(TextBox3.Text) >= 100 Then TextBox3.Text = "99"

        RadioBtn24.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        cRow = Trim(TextBox2.Text)
        cCol = Trim(TextBox3.Text)
        value = CInt(Trim(ComboBox2.Text))
        nDot = DotConvert(CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, CheckBox7.Checked, CheckBox8.Checked)
        interval = GetLEDStatus(Trim(ComboBox3.Text))
        ret = AB_3W_LB_DspNum(CurGwId, node, cRow, cCol, value, nDot, interval)
    End Sub

    Function CaluTagModeValue() As Integer
        Dim nMode As Integer = 0
        If TagModeChkBx1.Checked = True Then
            nMode += 1
        End If
        If TagModeChkBx2.Checked = True Then
            nMode += 2
        End If
        If TagModeChkBx3.Checked = True Then
            nMode += 4
        End If
        If TagModeChkBx4.Checked = True Then
            nMode += 8
        End If
        If TagModeChkBx5.Checked = True Then
            nMode += 16
        End If
        If TagModeChkBx6.Checked = True Then
            nMode += 32
        End If
        If TagModeChkBx7.Checked = True Then
            nMode += 64
        End If
        CaluTagModeValue = nMode
    End Function
    Function InitClearTags(ByVal nGw As Integer) As Integer
        Dim ret As Integer
        If lInitClearTags = False Then
            ret = AB_LB_DspStr(nGw, -252, 0, 0, -3)
            lInitClearTags = True
        End If
    End Function

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            lAutoScan = True
            Timer1.Enabled = True
        Else
            lAutoScan = False
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub BlinkAutoStatus()
        If lAutoScan = True Then
            If nShowReceiveSta = 6 Then
                If GroupBox5.Text = "○" & " Receive Data" Then
                    GroupBox5.Text = "●" & " Receive Data"
                Else
                    GroupBox5.Text = "○" & " Receive Data"
                End If
                nShowReceiveSta = 0
            End If
            nShowReceiveSta = nShowReceiveSta + 1
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim nNum As Integer
        nNum = 0
        Do While QwayGetDataB() > 0           'API COMMAND
            nNum = nNum + 1
        Loop
        ListBox2.Items.Add(TimeString & " Gateway = " & CurGwId & " " & ",Clear records :" & Trim(Str(nNum)))
        ListBox2.SelectedIndex = ListBox2.Items.Count - 1
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim ret As Integer
        ret = AB_GW_Status(CurGwId)
        ListBox2.Items.Add(TimeString & " Gateway = " & CurGwId & " " & ",Gateway status :" & AB_ErrMsg(ret))
        ListBox2.SelectedIndex = ListBox2.Items.Count - 1
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ret As Integer
        Button7.Text = "Please Wait!"
        Button7.Enabled = False
        ret = AB_GW_SetPollRang(CurGwId, -250)
        DelayItem = "SetBestPoll"
        Timer2.Tag = 2
        Timer2.Enabled = True
        Timer2.Interval = 8000 'if the time is not enough [AB_GW_SetPollRang(-250) -> need more time],it will not get the best polling range(it can't excute diangous too fast)
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim ret As Integer
        DelayItem = "reconnect"
        ret = ReConnect(CurGwId)
    End Sub

    Function ReConnect(ByVal nGwid As Integer) As Integer
        Dim ret As Integer, tmpfile As Object
        ret = AB_GW_Close(nGwid)             'Close the specified Gateway
        tmpfile = Shell(My.Application.Info.DirectoryPath & "\" & "darp.bat")  'clear ARP
        Timer2.Tag = 10
        Timer2.Enabled = True
        Timer2.Interval = 1000
    End Function
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

    Function DotConvert(ByVal ret6 As Boolean, ByVal ret5 As Boolean, ByVal ret4 As Boolean, ByVal ret3 As Boolean, ByVal ret2 As Boolean, ByVal ret1 As Boolean) As Integer
        Dim nNum As Integer
        nNum = 0
        If ret1 = True Then
            nNum = nNum + 32
        End If
        If ret2 = True Then
            nNum = nNum + 16
        End If
        If ret3 = True Then
            nNum = nNum + 8
        End If
        If ret4 = True Then
            nNum = nNum + 4
        End If
        If ret5 = True Then
            nNum = nNum + 2
        End If
        If ret6 = True Then
            nNum = nNum + 1
        End If
        DotConvert = nNum    'return dot parameters(decimal)
    End Function
    Function GetDelayTimes(ByVal nItem As Integer) As Single
        GetDelayTimes = nItem * 0.8
    End Function

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ret As Integer, node As Integer, nVolumn As Byte
        If Val(TextBox1.Text) >= 2 Then
            TextBox1.Text = Trim(Str(Val(TextBox1.Text) - 1))
        End If
        node = CInt(Trim(ComboBox1.Text))
        nVolumn = CInt(Trim(TextBox1.Text))
        ret = AB_Melody_Volume(CurGwId, node, nVolumn)
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ret As Integer, node As Integer, nVolumn As Byte
        If Val(TextBox1.Text) <= 15 Then
            TextBox1.Text = Trim(Str(Val(TextBox1.Text) + 1))
        End If
        node = CInt(Trim(ComboBox1.Text))
        nVolumn = CInt(Trim(TextBox1.Text))
        ret = AB_Melody_Volume(CurGwId, node, nVolumn)
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If RadioBtn24.Checked = True Then
            If Val(TextBox2.Text) >= 10 Then TextBox2.Text = "9"
        End If
        If Val(TextBox2.Text) >= 2 Then
            TextBox2.Text = Trim(Str(Val(TextBox2.Text) - 1))
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If RadioBtn24.Checked = True Then
            If Val(TextBox2.Text) <= 8 Then
                TextBox2.Text = Trim(Str(Val(TextBox2.Text) + 1))
            End If
        ElseIf RadioBtn25.Checked = True Then
            If Val(TextBox2.Text) <= 99998 Then
                TextBox2.Text = Trim(Str(Val(TextBox2.Text) + 1))
            End If
        End If
    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Val(TextBox3.Text) >= 2 Then
            TextBox3.Text = Trim(Str(Val(TextBox3.Text) - 1))
        End If
    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Val(TextBox3.Text) <= 98 Then
            TextBox3.Text = Trim(Str(Val(TextBox3.Text) + 1))
        End If
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox10.SelectedIndexChanged
        Dim ret As Integer, node As Integer, nSong As Byte, nBuzType As Byte
        RadioBtn21.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nSong = Val(Mid(ComboBox10.Text, 1, 1))
        nBuzType = Val(Mid(ComboBox6.Text, 1, 1))
        ret = AB_Melody_On(CurGwId, node, nSong, nBuzType)
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        Dim ret As Integer, node As Integer, nSong As Byte, nBuzType As Byte
        RadioBtn21.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        nSong = Val(Mid(ComboBox10.Text, 1, 1))
        nBuzType = Val(Mid(ComboBox6.Text, 1, 1))
        ret = AB_Melody_On(CurGwId, node, nSong, nBuzType)
    End Sub

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        Dim ret As Integer, node As Integer, batch As String, Loc As String, value As Integer, interval As Integer
        RadioBtn25.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        batch = Trim(TextBox2.Text)
        batch = Microsoft.VisualBasic.Right(Space(5) + Trim(batch), 5)
        Loc = Trim(TextBox3.Text)
        value = CInt(Trim(ComboBox2.Text))
        interval = GetLEDStatus(Trim(ComboBox3.Text))
        ret = AB_3W_CP_DspNum(CurGwId, node, batch, Loc, value, interval)
    End Sub

    Private Sub FeebackLabelColor()
        Label1.ForeColor = Color.Black
        Label2.ForeColor = Color.Black
        Label3.ForeColor = Color.Black
        Label4.ForeColor = Color.Black
        Label5.ForeColor = Color.Black
        Label6.ForeColor = Color.Black
        Label7.ForeColor = Color.Black
        Label8.ForeColor = Color.Black

        Label10.ForeColor = Color.Black
        Label11.ForeColor = Color.Black
        Label12.ForeColor = Color.Black
        Label13.ForeColor = Color.Black
        Label14.ForeColor = Color.Black
        Label15.ForeColor = Color.Black
        Label16.ForeColor = Color.Black
        Label17.ForeColor = Color.Black
        Label18.ForeColor = Color.Black
        Label19.ForeColor = Color.Black
        Label20.ForeColor = Color.Black
        Label21.ForeColor = Color.Black

        Label24.ForeColor = Color.Black
        Label25.ForeColor = Color.Black
        Label26.ForeColor = Color.Black
        Label27.ForeColor = Color.Black
        Label28.ForeColor = Color.Black

        GroupBox6.ForeColor = Color.Black
        GroupBox7.ForeColor = Color.Black
    End Sub

    Private Sub RadioBtn1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn1.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn3.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label3.ForeColor = Color.Red
        Label4.ForeColor = Color.Red
        GroupBox6.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn4.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn5.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label4.ForeColor = Color.Red
        Label5.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn6.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label5.ForeColor = Color.Red
        Label6.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn7.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label7.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn8.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn9.CheckedChanged
        FeebackLabelColor()
        Label8.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn10.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label10.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn11.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label11.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn12.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label12.ForeColor = Color.Red
        Label13.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn13.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label14.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn14.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        GroupBox7.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn15.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label15.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn16.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label16.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn17_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn17.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label3.ForeColor = Color.Red
        Label24.ForeColor = Color.Red
        Label25.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn18_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn18.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn19_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn19.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label26.ForeColor = Color.Red
        Label27.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn20_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn20.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label28.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn21_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn21.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label7.ForeColor = Color.Red
        Label17.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn22_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn22.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label18.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn23_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn23.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label2.ForeColor = Color.Red
        Label4.ForeColor = Color.Red
        Label19.ForeColor = Color.Red
        GroupBox6.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn24_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn24.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label2.ForeColor = Color.Red
        Label4.ForeColor = Color.Red
        Label20.ForeColor = Color.Red
        Label21.ForeColor = Color.Red
        GroupBox6.ForeColor = Color.Red
    End Sub

    Private Sub RadioBtn25_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioBtn25.CheckedChanged
        FeebackLabelColor()
        Label1.ForeColor = Color.Red
        Label2.ForeColor = Color.Red
        Label4.ForeColor = Color.Red
        Label20.ForeColor = Color.Red
        Label21.ForeColor = Color.Red
    End Sub
    Private Sub InitListBox1()
        Dim nGw As Integer
        If ListBox1.Items.Count = 0 Then
            For nGw = 1 To GwCnt
                ListBox1.Items.Add("Gw" & nGw & ":" & GwConf(nGw - 1).ip)
            Next nGw
        End If
    End Sub
    Private Sub DisplayTagMode()
        Dim nTagMode As Integer
        nTagMode = CaluTagModeValue()
        If nTagMode > 0 Then
            GroupBox7.Text = "Tag Mode(" & Trim(Str(nTagMode)) & ")"
        Else
            GroupBox7.Text = "Tag Mode"
        End If
    End Sub
    Private Sub TagModeChkBx1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TagModeChkBx1.CheckedChanged
        DisplayTagMode()
    End Sub

    Private Sub TagModeChkBx2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TagModeChkBx2.CheckedChanged
        DisplayTagMode()
    End Sub

    Private Sub TagModeChkBx3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TagModeChkBx3.CheckedChanged
        DisplayTagMode()
    End Sub

    Private Sub TagModeChkBx4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TagModeChkBx4.CheckedChanged
        DisplayTagMode()
    End Sub

    Private Sub TagModeChkBx5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TagModeChkBx5.CheckedChanged
        DisplayTagMode()
    End Sub

    Private Sub TagModeChkBx6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TagModeChkBx6.CheckedChanged
        DisplayTagMode()
    End Sub

    Private Sub TagModeChkBx7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TagModeChkBx7.CheckedChanged
        DisplayTagMode()
    End Sub

    Private Sub Button33_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        Dim ret As Integer, node As Integer, nVolumn As Byte
        If Val(TextBox1.Text) >= 2 Then
            TextBox1.Text = Trim(Str(Val(TextBox1.Text) - 1))
        End If
        node = CInt(Trim(ComboBox1.Text))
        nVolumn = CInt(Trim(TextBox1.Text))
        ret = AB_Melody_Volume(CurGwId, node, nVolumn)
    End Sub

    Private Sub Button34_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        Dim ret As Integer, node As Integer, nVolumn As Byte
        If Val(TextBox1.Text) <= 15 Then
            TextBox1.Text = Trim(Str(Val(TextBox1.Text) + 1))
        End If
        node = CInt(Trim(ComboBox1.Text))
        nVolumn = CInt(Trim(TextBox1.Text))
        ret = AB_Melody_Volume(CurGwId, node, nVolumn)
    End Sub

    Private Sub Button35_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        If RadioBtn24.Checked = True Then
            If Val(TextBox2.Text) >= 10 Then TextBox2.Text = "9"
        End If
        If Val(TextBox2.Text) >= 2 Then
            TextBox2.Text = Trim(Str(Val(TextBox2.Text) - 1))
        End If
    End Sub

    Private Sub Button12_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If RadioBtn24.Checked = True Then
            If Val(TextBox2.Text) <= 8 Then
                TextBox2.Text = Trim(Str(Val(TextBox2.Text) + 1))
            End If
        ElseIf RadioBtn25.Checked = True Then
            If Val(TextBox2.Text) <= 99998 Then
                TextBox2.Text = Trim(Str(Val(TextBox2.Text) + 1))
            End If
        End If
    End Sub

    Private Sub Button37_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        If Val(TextBox3.Text) >= 2 Then
            TextBox3.Text = Trim(Str(Val(TextBox3.Text) - 1))
        End If
    End Sub

    Private Sub Button36_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        If Val(TextBox3.Text) <= 98 Then
            TextBox3.Text = Trim(Str(Val(TextBox3.Text) + 1))
        End If
    End Sub

    Private Sub ComboBox13_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox13.SelectedIndexChanged

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged

    End Sub

    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        Dim node As Integer, ret As Integer, value As Integer, nDot, nLamp, nColor As Byte
        Dim nLed As Integer
        RadioBtn2.Checked = True
        node = CInt(Trim(ComboBox1.Text))
        value = CInt(Trim(ComboBox2.Text))
        nDot = DotConvert(CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, CheckBox7.Checked, CheckBox8.Checked)
        nLed = GetLEDStatus(Trim(ComboBox3.Text))
        nLamp = Val(Mid(ComboBox8.Text, 1, 1))
        nColor = Val(Mid(ComboBox7.Text, 1, 1))
        ret = AB_LB_DspNum(CurGwId, node, value, nDot, nColor)

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged

    End Sub
End Class
