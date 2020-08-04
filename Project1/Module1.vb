Option Strict Off
Option Explicit On
Imports System.Runtime.InteropServices
Imports System.Text
Module Module1
    Public GwCnt As Integer
    Public ScanTimer As Integer
    Public SysIniFile As String

    'Structure QWAY_CCB
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> Structure QWAY_CCB
        Dim ccblen As Short
        Dim ccbport As Byte
        Dim ccbdnode As Byte
        Dim ccbsnode As Byte
        Dim ccbcmd As Byte
        '<VBFixedArray(256)> Dim ccbdata() As Byte
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Public ccbdata As String

        'UPGRADE_TODO: 必須呼叫 "Initialize" 以初始化此結構的執行個體。 按一下以取得詳細資訊: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1026"'
        'Public Sub Initialize()
        '    ReDim ccbdata(256)
        'End Sub
    End Structure
    Sub LoadInitial()
        Dim ret As Integer

        GwCnt = AB_GW_Cnt()
        ReDim DiagGWTime(GwCnt)    'log gateway connecttimg time out
        ReDim lDiagGwOk(GwCnt)
        ret = AB_LoadConf()
    End Sub

    Public Class CQWAY_CCB
        Public ccblen As Short
        Public ccbport As Byte
        Public ccbdnode As Byte
        Public ccbsnode As Byte
        Public ccbcmd As Byte
        '<VBFixedArray(256)> Dim ccbdata() As Byte
        '<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Public ccbdata As String
        Public ccbdata(256) As Byte

        'UPGRADE_TODO: 必須呼叫 "Initialize" 以初始化此結構的執行個體。 按一下以取得詳細資訊: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1026"'
        'Public Sub Initialize()
        '    ReDim ccbdata(256)
        'End Sub
    End Class

    Structure CAPS_CCB
        Dim ccblen As Short
        Dim ccbport As Byte
        Dim ccbdnode As Byte
        Dim ccbsnode As Byte
        Dim ccbcmd As Byte
        Dim ccbsubcmd As Byte
        Dim ccbsubnode As Byte
        <VBFixedArray(256)> Dim ccbdata() As Byte

        'UPGRADE_TODO: 必須呼叫 "Initialize" 以初始化此結構的執行個體。 按一下以取得詳細資訊: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1026"'
        Public Sub Initialize()
            ReDim ccbdata(256)
        End Sub
    End Structure

    Structure LB_DSPSTR
        <VBFixedArray(7)> Dim dspstr() As Byte
        Public Sub Initialize()
            ReDim dspstr(7)
        End Sub
    End Structure

    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    '  // initial Gateway
    Declare Function AB_API_Open Lib "dapapi" () As Integer
    Declare Function AB_API_Close Lib "dapapi" () As Integer
    Declare Function AB_GW_Open Lib "dapapi" (ByVal Gateway_ID As Integer) As Integer
    Declare Function AB_GW_Close Lib "dapapi" (ByVal Gateway_ID As Integer) As Integer
    Declare Function AB_GW_Cnt Lib "dapapi" () As Integer
    Declare Function AB_GW_Conf Lib "dapapi" (ByVal ndx As Integer, ByRef Gateway_ID As Integer, ByRef ip As Byte, ByRef port As Integer) As Integer
    Declare Function AB_GW_Ndx2ID Lib "dapapi" (ByVal ndx As Long) As Integer
    Declare Function AB_GW_ID2Ndx Lib "dapapi" (ByVal Gateway_ID As Long) As Integer
    Declare Function AB_GW_InsConf Lib "dapapi" (ByVal Gateway_ID As Long, ByVal ip As Byte, ByVal port As Long) As Integer
    Declare Function AB_GW_UpdConf Lib "dapapi" (ByVal Gateway_ID As Long, ByVal ip As Byte, ByVal port As Long) As Integer
    Declare Function AB_GW_DelConf Lib "dapapi" (ByVal Gateway_ID As Long) As Integer

    '  // Get/Send message from/to Gateway
    Declare Function AB_GW_RcvMsg Lib "dapapi" (ByVal Gateway_ID As Long, ByVal ccb As QWAY_CCB) As Integer
    Declare Function AB_GW_SndMsg Lib "dapapi" (ByVal Gateway_ID As Long, ByVal ccb As QWAY_CCB) As Integer
    Declare Function AB_GW_RcvReady Lib "dapapi" (ByVal Gateway_ID As Long) As Integer
    Declare Function AB_GW_RcvButton Lib "dapapi" (ByRef data As Byte, ByRef data_cnt As Integer) As Integer

    Declare Function AB_CLTAG_SetAddr Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal Mode As Short) As Short
    Declare Function AB_CLTAG_DspAddr Lib "dapapi" (ByVal port As Short, ByVal Mode As Short) As Short

    '  // Get Gateway status
    Declare Function AB_GW_Status Lib "dapapi" (ByVal Gateway_ID As Integer) As Integer
    Declare Function AB_GW_AllStatus Lib "dapapi" (ByVal status As Byte) As Integer
    Declare Function AB_GW_TagDiag Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal port As Long) As Integer
    Declare Function AB_GW_SetPollRang Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal poll_range As Integer) As Integer

    '  // Send Message to picking tag
    Declare Function AB_LB_DspNum Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Disp_Int As Integer, ByVal Dot As Byte, ByVal interval As Integer) As Integer
    Declare Function AB_LB_DspStr Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Disp_Str As String, ByVal Dot As Byte, ByVal interval As Integer) As Integer
    Declare Function AB_LB_DspStr2 Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Disp_Str As Byte, ByVal interval As Long) As Integer
    Declare Function AB_LB_DspOff Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer) As Integer
    Declare Function AB_LB_DspAddr Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer) As Integer
    Declare Function AB_LB_LedOn Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Lamp_STA As Byte, ByVal interval As Integer) As Integer
    Declare Function AB_LB_BuzOn Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Buzzer_Type As Byte) As Integer
    Declare Function AB_LB_SetMode Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Pick_Mode As Byte) As Integer
    Declare Function AB_LB_Simulate Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Simulate_Mode As Byte) As Integer
    Declare Function AB_LB_SetLock Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Lock_State As Byte, ByVal Lock_key As Byte) As Integer
    Declare Function AB_LED_Status Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Lamp_COLOR As Byte, ByVal Lamp_STA As Byte) As Long
    Declare Function AB_BUZ_On Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Buzzer_Type As Byte) As Long

    Declare Function AB_Tag_RcvMsg Lib "dapapi" (ByRef Gateway_ID As Integer, ByRef node_addr As Integer, ByRef subcmd As Short, ByRef msg_type As Short, ByRef data As Byte, ByRef data_cnt As Integer) As Integer
    Declare Function AB_TAG_Reset Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer) As Integer
    Declare Function AB_Tag_ChgAddr Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal new_tag As Long) As Integer

    Declare Function AB_LED_Dsp Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Lamp_STA As Byte, ByVal interval As Short) As Integer
    Declare Function AB_TAG_mode Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Save_Mode As Integer, ByVal Mode_Type As Byte) As Integer
    Declare Function AB_TAG_ButtonDelay Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal DelayTime As Single) As Short
    Declare Function AB_TAG_CountDigit Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Digit As Short) As Short
    Declare Function AB_TAG_Complete Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Complete_State As Byte) As Integer

    '  // Send Message to AT503A & AT502V
    Declare Function AB_AV_LB_DspNum Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Disp_Int As Long, ByVal Arrow As Byte, ByVal Dot As Byte, ByVal interval As Integer) As Integer
    Declare Function AB_3W_LB_DspNum Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal FRow As String, ByVal Ecolumn As String, ByVal Disp_Int As Integer, ByVal Dot As Integer, ByVal interval As Integer) As Integer
    Declare Function AB_Melody_On Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal song As Byte, ByVal Buzzer_Type As Byte) As Integer
    Declare Function AB_Melody_Volume Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Vol As Byte) As Integer

    '  // Send meaasge to special tag level devices
    Declare Function AB_CCD_Action Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal Action As Integer) As Integer
    Declare Function AB_KBN_DspMsg Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal dsp_mode As Integer, ByRef msg() As Byte) As Integer

    '  //-----MMI-19
    Declare Function AB_DCS_InputMode Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal input_mode As Byte) As Integer
    Declare Function AB_DCS_BufSize Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal buf_size As Byte) As Integer
    Declare Function AB_DCS_SetConf Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal enable_status As Byte, ByVal disable_status As Byte) As Integer
    Declare Function AB_DCS_ReqConf Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DCS_GetVer Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DCS_Reset Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DCS_SetRows Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal rows As Byte) As Integer
    Declare Function AB_DCS_SimulateKey Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal key_code As Byte) As Integer
    Declare Function AB_DCS_Cls Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DCS_Buzzer Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal alarm_time As Byte, ByVal alarm_cnt As Byte) As Integer
    Declare Function AB_DCS_ScrollUp Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal up_rows As Byte) As Integer
    Declare Function AB_DCS_ScrollDown Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal down_rows As Byte) As Integer
    Declare Function AB_DCS_ScrollHome Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DCS_ScrollEnd Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DCS_SetCursor Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal row As Byte, ByVal column As Byte) As Integer
    Declare Function AB_DCS_DspStrE Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal dsp_str As Byte, ByVal dsp_cnt As Long) As Integer
    Declare Function AB_DCS_DspStrC Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal dsp_str As Byte, ByVal dsp_cnt As Long) As Integer

    '  //-----DT200
    Declare Function AB_DT2_DspStr Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal dsp_str As Byte, ByVal dsp_cnt As Long) As Integer
    Declare Function AB_DT2_EnableCounter Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DT2_DisableCounter Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DT2_ReadCounter Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer) As Integer
    Declare Function AB_DT2_SetCounter Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal count As Long) As Integer
    Declare Function AB_DT2_Buzzer Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal alarm_time As Byte, ByVal alarm_cnt As Byte) As Integer
    Declare Function AB_DT2_AutoReport Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal enable_flag As Byte) As Integer
    Declare Function AB_DT2_FunKeyConfig Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal config As Byte) As Integer
    Declare Function AB_DT2_Terminator Lib "dapapi" (ByVal Gateway_ID As Long, ByVal node_addr As Integer, ByVal config As Byte) As Integer

    '  // Send Message to 50C picking tag
    Declare Function AB_AHA_DspStr Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Disp_Str As String, ByVal BeConfirm As Byte, ByVal interval As Byte) As Integer
    Declare Function AB_AHA_ReDsp Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer) As Integer
    Declare Function AB_AHA_ClrDsp Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer) As Integer
    Declare Function AB_AHA_LED_Dsp Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Lamp_Type As Byte, ByVal Lamp_STA As Integer) As Integer
    Declare Function AB_AHA_BUZ_On Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Buz_STA As Integer) As Integer

    ' // -----50A(5-2-3)
    Declare Function AB_3W_CP_DspNum Lib "dapapi" (ByVal Gateway_ID As Integer, ByVal node_addr As Integer, ByVal Lot As String, ByVal Loc_Renamed As String, ByVal Disp_Int As Integer, ByVal interval As Integer) As Integer

    Function AB_ErrMsg(ByRef ret As Short) As String
        Dim tmpstr As Object
        Select Case ret
            Case -3
                tmpstr = "Parameter data is error !"
            Case -2
                tmpstr = "TCP is not created yet !"
            Case -1
                tmpstr = "DAP_ID out of range !"
            Case 0
                tmpstr = "Closed"
            Case 1
                tmpstr = "Open"
            Case 2
                tmpstr = "Listening"
            Case 3
                tmpstr = "Connection is Pending"
            Case 4
                tmpstr = "Resolving the host name"
            Case 5
                tmpstr = "Host is Resolved"
            Case 6
                tmpstr = "Waitting"  '"Waiting to Connect"
            Case 7
                tmpstr = "ok"        '"Connected ok "
            Case 8
                tmpstr = "Connection is closing"
            Case 9
                tmpstr = "State error has occurred"
            Case 10
                tmpstr = "Connection state is undetermined"
            Case Else
                tmpstr = "Unknown Error Code"
        End Select

        AB_ErrMsg = tmpstr
    End Function
        Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, _
                ByVal lpKeyName As String, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As StringBuilder, _
                ByVal nSize As Integer, _
                ByVal lpFileName As String) As Integer

        Sub baca_text()

            Dim res As Integer
            Dim sb As StringBuilder

            sb = New StringBuilder(500)
            res = GetPrivateProfileString("contoh", "server", "", sb, sb.Capacity, My.Application.Info.DirectoryPath & "baca.txt")
            Console.WriteLine("GetPrivateProfileStrng returned : " & res.ToString())
            Console.WriteLine("server=" & sb.ToString())
            Console.ReadLine()

        End Sub
End Module