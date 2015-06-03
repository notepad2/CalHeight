Imports System
Imports System.IO
Imports System.Text
Public Class Form1

    Dim r, t, zxzh, zxbg, qpd, hpd, ypj, hp As Double
    Dim dunhao, hxgs, i, j, hangshu As Integer
    Dim array1 As Array

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        r = Val(TextBox1.Text)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        t = Val(TextBox2.Text)
        TextBox7.Text = Cale(r, t)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        zxzh = Val(TextBox3.Text)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        zxbg = Val(TextBox4.Text)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        qpd = Val(TextBox5.Text)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        hpd = Val(TextBox6.Text)
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        ypj = Val(TextBox8.Text)
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged, TextBox9.TextChanged
        hp = Val(TextBox9.Text)
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        dunhao = Val(TextBox10.Text)
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        hxgs = Val(TextBox11.Text)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ReGen(dunhao, hxgs)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        '计算桩号
        Dim i, j As Integer
        Dim dis, jszh, height As Double

        For i = 0 To dunhao - 1 Step 1
            For j = 0 To hxgs - 1 Step 1

                dis = DataGridView1.Item(2, j + i * hxgs).Value

                If j <> hxgs \ 2 Then
                    jszh = CalZh(DataGridView1.Item(1, i * hxgs + hxgs \ 2).Value, ypj * Math.PI / 180, dis)
                    DataGridView1.Item(1, j + i * hxgs).Value = jszh
                    DataGridView1.Item(1, j + i * hxgs).ReadOnly = True
                Else

                End If
            Next j
        Next i


        For i = 0 To hangshu - 1 Step 1
            jszh = DataGridView1.Item(1, i).Value


            If DataGridView1.Item(1, i).FormattedValue = Nothing Or DataGridView1.Item(2, i).FormattedValue = Nothing Then
                DataGridView1.Item(3, i).Value = Nothing
                DataGridView1.Item(4, i).Value = Nothing
            Else
                '计算中心线位置标高
                DataGridView1.Item(3, i).Value = CalHeight(r, t, zxzh, zxbg, qpd / 100, hpd / 100, jszh)
                DataGridView1.Item(3, i).ReadOnly = True
                height = CalHeight(r, t, zxzh, zxbg, qpd / 100, hpd / 100, jszh)
                '计算点位置标高
                dis = DataGridView1.Item(2, i).Value
                DataGridView1.Item(4, i).Value = CalDianHei(height, dis, hp / 100)
                DataGridView1.Item(4, i).ReadOnly = True
            End If
        Next i

        'Array.Copy(SaveToArray(hangshu, 5, DataGridView1), array1, 0)
        'array1 = SaveToArray(hangshu, 5, DataGridView1)
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '公用函数部分

    '计算e值函数
    Public Function Cale(ByVal r As Double, ByVal t As Double) As Double
        Cale = Math.Round(Math.Sqrt(r * r + t * t) - r, 3)
    End Function
    '重新生成计算表格函数
    Public Function ReGen(ByVal dh As Integer, ByVal hx As Integer)

        Dim i, j As Integer

        '重置所有单元格内容为空
        DataGridView1.Rows.Clear()

        hangshu = dh * hx
        DataGridView1.RowCount = hangshu

        For i = 0 To dh - 1 Step 1
            For j = 0 To hx - 1 Step 1
                DataGridView1.Item(0, j + i * hx).Value = i
                DataGridView1.Item(0, j + i * hx).ReadOnly = True

                If j = hx \ 2 Then
                    DataGridView1.Item(1, j + i * hx).Style.BackColor = Color.Green
                    DataGridView1.Item(2, j + i * hx).Value = 0
                    DataGridView1.Item(2, j + i * hx).ReadOnly = True
                Else
                    DataGridView1.Item(2, j + i * hx).Style.BackColor = Color.Green
                    DataGridView1.Item(1, j + i * hx).ReadOnly = True
                    DataGridView1.Item(3, j + i * hx).ReadOnly = True
                    DataGridView1.Item(4, j + i * hx).ReadOnly = True

                End If
            Next j
        Next i
        ReGen = 0
    End Function

    '桩号计算函数
    Public Function CalZh(ByVal zxzh As Double, ByVal ypj As Double, ByVal dis As Double) As Double
        CalZh = zxzh - dis * Math.Tan(Math.PI / 2 - ypj)
    End Function

    '中心线位置标高计算函数
    Public Function CalHeight(ByVal r As Double, ByVal t As Double, ByVal zxzh As Double, ByVal zxbg As Double,
                      ByVal qpd As Double, ByVal hpd As Double, ByVal jszh As Double) As Double
        Dim e As Double
        Dim flag As Integer

        e = Math.Sqrt(r * r + t * t) - r

        If (qpd > hpd) Then
            flag = 1
        Else
            flag = -1
        End If

        If jszh < zxzh - t Then
            CalHeight = zxbg + (jszh - zxzh) * qpd
        ElseIf jszh > zxzh + t Then
            CalHeight = zxbg + (jszh - zxzh) * hpd
        Else
            If (jszh < zxzh) Then
                CalHeight = Math.Round(zxbg - qpd * (zxzh - jszh) - flag * Math.Pow(t - Math.Abs(zxzh - jszh), 2) / 2 / r, 3)
            Else
                CalHeight = Math.Round(zxbg - hpd * (zxzh - jszh) - flag * Math.Pow(t - Math.Abs(zxzh - jszh), 2) / 2 / r, 3)
            End If
        End If
    End Function

    '点位置标高计算函数
    Public Function CalDianHei(ByVal height As Double, ByVal dis As Double, ByVal hp As Double) As Double
        If dis > 0 Then
            CalDianHei = height - dis * hp
        Else
            CalDianHei = height + dis * hp
        End If
    End Function

    '将表格数值保存到数组
    'Public Function SaveToArray(ByVal a As Integer, ByVal b As Integer, ByRef dgv1 As DataGridView) As Array
    'Dim i, j As Integer
    'Dim ary1(a, b) As Double
    'For i = 0 To a - 1 Step 1
    'For j = 0 To b - 1 Step 1
    'ary1(i, j) = Val(dgv1.Item(j, i).Value)
    'Next j
    'Next i
    'SaveToArray = ary1
    'End Function

    '文件保存函数
    Public Function WriteFile(ByVal filename As String)

        'Dim file As StreamWriter
        'file = My.Computer.FileSystem.OpenTextFileWriter(filename, False)
        Dim i, j As Integer
        Dim file As New StreamWriter(filename, False)

        '写入各主要参数
        file.Write(r)
        file.Write(" ")
        file.Write(t)
        file.Write(" ")
        file.Write(zxzh)
        file.Write(" ")
        file.Write(zxbg)
        file.Write(" ")
        file.Write(qpd)
        file.Write(" ")
        file.Write(hpd)
        file.Write(" ")
        file.Write(ypj)
        file.Write(" ")
        file.Write(hp)
        file.Write(" ")
        file.Write(dunhao)
        file.Write(" ")
        file.WriteLine(hxgs)

        '写入表格数据
        For i = 0 To hangshu - 1 Step 1
            For j = 0 To 4

                '按矩阵形式写入表格
                If j <> 4 Then
                    If DataGridView1.Item(j, i).FormattedValue = Nothing Then
                        file.Write("\")
                        file.Write(" ")
                    Else
                        file.Write(DataGridView1.Item(j, i).Value)
                        file.Write(" ")
                    End If
                Else
                    If DataGridView1.Item(j, i).FormattedValue = Nothing Then
                        file.WriteLine("\")
                    Else
                        file.WriteLine(DataGridView1.Item(j, i).Value)
                    End If
                End If
            Next j
        Next i

        file.Close()
        WriteFile = 0

    End Function

    '读取并使用文件函数
    Public Function UseFile(ByVal filename As String)
        Dim sr As New StreamReader(filename)
        Dim str() As String
        Dim e As Double

        str = Split(sr.ReadLine, " ")

        r = Val(str(0))
        TextBox1.Text = r

        t = Val(str(1))
        TextBox2.Text = t

        zxzh = Val(str(2))
        TextBox3.Text = zxzh

        zxbg = Val(str(3))
        TextBox4.Text = zxbg

        qpd = Val(str(4))
        TextBox5.Text = qpd

        hpd = Val(str(5))
        TextBox6.Text = hpd

        e = Cale(r, t)
        TextBox7.Text = e

        ypj = Val(str(6))
        TextBox8.Text = ypj

        hp = Val(str(7))
        TextBox9.Text = hp

        dunhao = Val(str(8))
        TextBox10.Text = dunhao

        hxgs = Val(str(9))
        TextBox11.Text = hxgs

        ReGen(dunhao, hxgs)

        hangshu = dunhao * hxgs

        For i = 0 To hangshu - 1 Step 1
            str = Split(sr.ReadLine, " ")
            For j = 0 To 4 Step 1
                If str(j) = "\" Then
                    DataGridView1.Item(j, i).Value = Nothing
                Else
                    DataGridView1.Item(j, i).Value = Val(str(j))
                End If
            Next j
        Next i

        sr.Close()
        UseFile = 0

    End Function

    '重置所有内容
    Public Function ResetAll()
        TextBox1.Text = Nothing
        TextBox2.Text = Nothing
        TextBox3.Text = Nothing
        TextBox4.Text = Nothing
        TextBox5.Text = Nothing
        TextBox6.Text = Nothing
        TextBox7.Text = Nothing
        TextBox8.Text = Nothing
        TextBox9.Text = Nothing
        TextBox10.Text = Nothing
        TextBox11.Text = Nothing

        ReGen(0, 0)

        ResetAll = 0
    End Function
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '菜单功能部分


    '新建按钮
    Private Sub 新建ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 新建ToolStripMenuItem.Click
        ResetAll()
    End Sub

    '保存按钮
    Private Sub 保存ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 保存ToolStripMenuItem.Click
        Dim SaveFileDialog1 As New SaveFileDialog
        SaveFileDialog1.Filter = "标高计算文件|*.bgjs|所有文件|*.*"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            WriteFile(SaveFileDialog1.FileName)
        End If
    End Sub

    '打开按钮
    Private Sub 打开ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 打开ToolStripMenuItem.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        OpenFileDialog1.Filter = "标高计算文件|*.bgjs|所有文件|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            UseFile(OpenFileDialog1.FileName)

        End If
    End Sub

    '退出按钮
    Private Sub 退出ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 退出ToolStripMenuItem.Click
        Close()
    End Sub

End Class

'待实现功能
'1、文件打开保存操作
'2、表格内粘贴
'3、帮助及提示

'20150323更新0.3版本，菜单项打开、保存、新建、退出已经初步实现。
'20150323更新0.31版本，降低.net框架版本为2.0
'20150603更新0.32版本，修正道路中心线上的点标高算法错误。