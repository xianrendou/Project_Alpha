Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Web
Imports System.Xml

Public Class MainForm
    Dim tmpTrans As String
    Dim wordop As New Word.Application
    Dim worddo As Word.Document
    Dim txtfile As String
    Dim reader As System.IO.StreamReader
    Dim currentword As String
    Dim CMeaning As String
    Dim currentPro As String

    Function Search(tmpWord As String) As String
        'http://dict.youdao.com/search?q=单词&keyfrom=dict.index
        Dim XH As Object
        Dim s() As String
        Dim str_tmp As String
        Dim str_base As String
        Dim tmpPhonetic As String
        Dim yb As String

        tmpTrans = ""
        tmpPhonetic = ""

        '开启网页
        XH = CreateObject("Microsoft.XMLHTTP")
        On Error Resume Next
        XH.Open("get", "http://dict.youdao.com/search?q=" & tmpWord & "&keyfrom=dict.index", False)
        XH.send()
        On Error Resume Next
        str_base = XH.responseText
        XH.Close()
        XH = Nothing

        yb = Split(Split(str_base, "<div id=""webTrans"" class=""trans-wrapper trans-tab"">")(0), "<span class=""keyword"">")(1)

        '取音标
        If UBound(Split(yb, "<span class=""pronounce"">美")) = 1 Then
            '美式音标
            tmpPhonetic = Split((Split(Split(yb, "<span class=""pronounce"">美")(1), "<span class=""phonetic"">")(1)), "</span>")(0)
            On Error Resume Next
        Else
            tmpPhonetic = Split((Split(yb, "<span class=""phonetic"">")(1)), "</span>")(0)
            On Error Resume Next
        End If

        '取中文翻译
        str_tmp = Split((Split(yb, "<div class=""trans-container"">")(1)), "</div>")(0)
        str_tmp = Split((Split(str_tmp, "<ul>")(1)), "</ul>")(0)
        s = Split(str_tmp, "<li>")
        tmpTrans = Split(s(LBound(s) + 1), "</li")(0)
        For i = LBound(s) + 2 To UBound(s)
            tmpTrans = tmpTrans & Chr(10) & Split(s(i), "</li")(0)
        Next
        Return tmpPhonetic
    End Function

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        End
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Try
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                txtfile = OpenFileDialog1.FileName
                Call readfile()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DialogStyle.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, iCount As Integer
        iCount = ListBox1.Items.Count()
        For i = 0 To iCount - 1
            currentword = ListBox1.Items.Item(i).ToString
            currentPro = Search(currentword)
            wordop.Selection.TypeText(currentword & " " & currentPro & " " & tmpTrans & vbCrLf)
        Next
        wordop.Visible = True
    End Sub

    Sub readfile()
        reader = New IO.StreamReader(txtfile)
        While Not reader.EndOfStream
            ListBox1.Items.Add(reader.ReadLine)
        End While
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        worddo = wordop.Documents.Add()
        Config_File()
    End Sub

    Private Sub Config_File()
        Dim config = New XmlDocument()
        config.Load("config.xml") '读取XML文档
        Dim fr = config.SelectSingleNode("config").SelectSingleNode("firstrun").InnerText '判断首次运行
        If fr = "true" Then
            firstRun.ShowDialog()
        End If
    End Sub
End Class
