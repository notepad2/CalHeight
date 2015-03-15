Public Class Form1

    Dim r, t, zxzh, zxbg, qpd, hpd, jszh, jsbg, ypj, hp As Double
    Dim dunhao, hxgs, i, j, hangshu As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Function Cale(ByVal r As Double, ByVal t As Double) As Double
        Cale = Math.Round(Math.Sqrt(r * r + t * t) - r, 3)
    End Function


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
        qpd = Val(TextBox5.Text) / 100
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        hpd = Val(TextBox6.Text) / 100
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        ypj = Val(TextBox8.Text) * Math.PI / 180
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged, TextBox9.TextChanged
        hp = Val(TextBox9.Text) / 100
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

                If j <> CInt(Fix(hxgs / 2)) Then
                    jszh = CalZh(DataGridView1.Item(1, i * hxgs + CInt(Fix(hxgs / 2))).Value, ypj, dis)
                    DataGridView1.Item(1, j + i * hxgs).Value = jszh
                    DataGridView1.Item(1, j + i * hxgs).ReadOnly = True
                Else

                End If
            Next j
        Next i


        For i = 0 To hangshu - 1 Step 1
            jszh = DataGridView1.Item(1, i).Value

            '计算中心线位置标高
            DataGridView1.Item(3, i).Value = CalHeight(r, t, zxzh, zxbg, qpd, hpd, jszh)
            height = CalHeight(r, t, zxzh, zxbg, qpd, hpd, jszh)
            DataGridView1.Item(3, i).ReadOnly = True

            '计算点位置标高
            dis = DataGridView1.Item(2, i).Value
            DataGridView1.Item(4, i).Value = CalDianHei(height, dis, hp)
            DataGridView1.Item(4, i).ReadOnly = True
        Next i

    End Sub

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded

    End Sub

    '重新生成计算表格函数
    Public Function ReGen(ByVal dh As Integer, ByVal hx As Integer)

        Dim i, j As Integer

        '重置所有单元格内容为空
        DataGridView1.Rows.Clear()

        hangshu = dh * hx
        DataGridView1.RowCount = hangshu

        For i = 0 To dunhao - 1 Step 1
            For j = 0 To hxgs - 1 Step 1
                DataGridView1.Item(0, j + i * hx).Value = i
            Next j
            DataGridView1.Item(2, CInt(Fix(hx / 2 + i * hx))).Value = 0
            DataGridView1.Item(2, CInt(Fix(hx / 2 + i * hx))).ReadOnly = True

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

        e = Math.Sqrt(r * r + t * t) - r

        If jszh < zxzh - t Then
            CalHeight = zxbg + (jszh - zxzh) * qpd
        ElseIf jszh > zxzh + t Then
            CalHeight = zxbg + (jszh - zxzh) * hpd
        Else
            If (qpd >= 0) Then
                CalHeight = Math.Round(zxbg - e - (r - Math.Sqrt(r * r - (zxzh - jszh) * (zxzh - jszh))), 3)
            Else
                CalHeight = Math.Round(zxbg + e + (r - Math.Sqrt(r * r - (zxzh - jszh) * (zxzh - jszh))), 3)
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
End Class

'待实现功能
'1、文件打开保存操作
'2、表格内粘贴
'3、帮助及提示