Imports System.Xml

Public Class firstRun
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim config = New XmlDocument()
        config.Load("config.xml") '读取XML文档
        config.SelectSingleNode("config").SelectSingleNode("firstrun").InnerText = "false" '修改首次运行值
        config.Save("config.xml")
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("iexplore.exe", "https://github.com/xianrendou/Project_Alpha")
    End Sub
End Class