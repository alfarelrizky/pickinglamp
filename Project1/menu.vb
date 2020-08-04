Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.Odbc
Public Class menu_form
    Inherits System.Windows.Forms.Form
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

    'publikasi
    'pos1
    Public liffting1(8) As String
    Public vin1(8) As String
    Public suffix1(8) As String
    Public model1(8) As String
    Public color1(8) As String
    Public entry(8) As String
    'pos 1 array partname
    Public part_name(300) As String
    Public jum_part1 As Integer

    Private Sub menu_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Me.Close()
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs)
    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        testlamp.Show()
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        monitoring.Show()
    End Sub

    Private Sub ToolStripStatusLabel1_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel1.Click

    End Sub
End Class