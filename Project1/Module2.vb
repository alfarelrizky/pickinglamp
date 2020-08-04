Option Strict Off
Option Explicit On
Module Module2

    Sub MsgBox2(ByRef msg As String)
        Dim lst As ListBox
        Dim timestr As String

        lst = Form1.DefInstance.ListBox2
        timestr = TimeString

        If lst.Items.Count >= 100 Then
            lst.Items.RemoveAt(0)
        End If
        lst.Items.Add(timestr & " " & msg)
        lst.SelectedIndex = lst.Items.Count - 1
    End Sub

    Function HexIp(ByRef ip_str As String, ByRef op_str As String) As Short
        Dim start, cnt As Short
        Dim tmpstr As String

        op_str = ""
        start = 1
        cnt = 0
        Do
            tmpstr = Mid(ip_str, start, 1)
            start = start + 1
            If tmpstr = "" Then Exit Do
            If tmpstr = "\" Then
                tmpstr = "&H"
                tmpstr = tmpstr & Mid(ip_str, start, 2)
                start = start + 2
                tmpstr = Chr(Val(tmpstr))
            End If
            op_str = op_str & tmpstr
            cnt = cnt + 1
        Loop
        HexIp = cnt
    End Function

    Function HexIpB(ByRef ip_str As String, ByRef op_str() As Byte, ByRef op_start As Short) As Short
        Dim start, cnt As Short
        Dim tmpstr As String
        Dim data As Short

        start = 1
        cnt = 0 + op_start
        Do
            tmpstr = Mid(ip_str, start, 1)
            start = start + 1
            If tmpstr = "" Then Exit Do
            If tmpstr = "\" Then
                tmpstr = "&H" & Mid(ip_str, start, 2)
                start = start + 2
                op_str(cnt) = Val(tmpstr)
                cnt = cnt + 1
            Else
                data = Asc(tmpstr)
                op_str(cnt) = (data Mod 256) And &HFFS
                cnt = cnt + 1
                If data And &HFF00S Then
                    op_str(cnt) = (data \ 256) And &HFFS
                    cnt = cnt + 1
                End If
            End If
        Loop
        HexIpB = cnt

    End Function


    Function HexOp0(ByRef ip_str As Object, ByRef cnt As Short, ByRef op_str As String) As Short
        Dim start As Short
        Dim tmpstr As String
        Dim ch As Short

        op_str = ""
        For start = 1 To cnt
            'UPGRADE_WARNING: 無法解析物件 ip_str 的預設屬性。 按一下以取得詳細資訊: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
            ch = Asc(Mid(ip_str, start, 1))
            'If ch = 0 Or ch = &H5C Or ch > &H80 Then
            If ch < &H20S Or ch = &H5CS Or ch >= &H80S Then
                tmpstr = Hex(ch)
                If Len(tmpstr) < 2 Then
                    op_str = op_str & "\0" & tmpstr
                Else
                    op_str = op_str & "\" & tmpstr
                End If
            Else
                op_str = op_str & Chr(ch)
            End If
        Next start
    End Function

    Function HexOp2(ByRef ip_str As Object, ByRef cnt As Short) As String
        Dim ret As Short
        Dim op_str As String

        If cnt < 0 Then cnt = Len(ip_str)

        ret = HexOp0(ip_str, cnt, op_str)
        HexOp2 = op_str
    End Function

    Function ldump(ByRef ip_str As Object, ByRef op_str As Object) As Short
        Dim start As Short
        Dim tmpstr As String

        'UPGRADE_WARNING: 無法解析物件 op_str 的預設屬性。 按一下以取得詳細資訊: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
        op_str = ""
        start = 1
        Do
            'UPGRADE_WARNING: 無法解析物件 ip_str 的預設屬性。 按一下以取得詳細資訊: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
            tmpstr = Mid(ip_str, start, 1)
            start = start + 1
            If tmpstr = "" Then Exit Do
            tmpstr = Hex(Asc(tmpstr)) & " "
            If Len(tmpstr) < 3 Then
                tmpstr = "0" & tmpstr
            End If
            'UPGRADE_WARNING: 無法解析物件 op_str 的預設屬性。 按一下以取得詳細資訊: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
            op_str = op_str + tmpstr
        Loop

    End Function

    Function ldump2(ByRef ip_str As String, ByRef cnt As Short) As String
        Dim start As Short
        Dim tmpstr, op_str As String

        op_str = ""
        If cnt < 0 Then cnt = Len(ip_str)
        start = 1
        Do
            tmpstr = Mid(ip_str, start, 1)
            start = start + 1
            If tmpstr = "" Then Exit Do
            tmpstr = Hex(Asc(tmpstr)) & " "
            If Len(tmpstr) < 3 Then
                tmpstr = "0" & tmpstr
            End If
            op_str = op_str & tmpstr

            If start > cnt Then Exit Do
        Loop

        ldump2 = op_str
    End Function

    Function ldump3(ByRef ip_str() As Byte, ByRef start As Short, ByRef cnt As Short) As String
        Dim i As Short
        Dim tmpstr, op_str As String

        op_str = ""
        For i = 0 To cnt - 1
            tmpstr = Hex(ip_str(start + i)) & " "
            If Len(tmpstr) < 3 Then
                tmpstr = "0" & tmpstr
            End If
            op_str = op_str & tmpstr
        Next i

        ldump3 = op_str
    End Function
    'Function CreateBitStr(ByRef nAsc As Byte) As String
    '    Dim cBitStr As String
    '    Dim aBinary(8) As Integer
    '    Dim lmore As Boolean
    '    Dim aBitAry(8) As Boolean
    '    Dim I As Integer

    '    aBinary(1) = 1
    '    aBinary(2) = 2
    '    aBinary(3) = 4
    '    aBinary(4) = 8
    '    aBinary(5) = 16
    '    aBinary(6) = 32
    '    aBinary(7) = 64
    '    aBinary(8) = 128
    '    lmore = True
    '    Do While lmore = True
    '        If nAsc = 0 Then
    '            lmore = False
    '        Else
    '            For I = 1 To 8
    '                If nAsc >= aBinary(8 - I + 1) Then
    '                    nAsc = nAsc - aBinary(8 - I + 1)
    '                    aBitAry(8 - I + 1) = True
    '                    Exit Do
    '                End If
    '            Next I
    '        End If
    '    Loop
    '    cBitStr = ""
    '    For I = 1 To 8
    '        If aBitAry(I) = True Then
    '            cBitStr = cBitStr & "1"
    '        Else
    '            cBitStr = cBitStr & "0"
    '        End If
    '    Next I
    '    CreateBitStr = cBitStr
    'End Function

    Function str2hex(ByRef tmpstr As Object) As String
        '// change string to Hex format string (Double-Byte Code)
        Dim cnt, i As Short
        Dim tmpstr2 As String
        Dim data As Short

        cnt = Len(tmpstr)
        tmpstr2 = ""
        For i = 1 To cnt
            'UPGRADE_WARNING: 無法解析物件 tmpstr 的預設屬性。 按一下以取得詳細資訊: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1037"'
            data = Asc(Mid(tmpstr, i, 1)) 'double byte code
            If (data And &HFF00S) = 0 Then
                tmpstr2 = tmpstr2 & Mid(Hex(&H100S + data), 2)
            Else
                tmpstr2 = tmpstr2 & Mid(Hex(&H10000 + CShort(data And &HFFFF)), 2)
            End If
        Next i
        str2hex = tmpstr2
    End Function
End Module