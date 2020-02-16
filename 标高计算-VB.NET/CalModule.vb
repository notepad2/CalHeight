Imports System.IO
Module CalModule

    '按桩号计算道路中心线标高模块，输入竖曲线参数表（list嵌套array格式）、以及计算桩号，输出计算所得标高。
    Public Function CalBySQX(ByVal BGS As List(Of Array), jszh As Double) As Double
        Dim i, j, maxrow, maxcol, row, flag As Integer
        Dim zxzh, zxbg, r, t, e, qpd, hpd As Double

        i = 0
        j = 0
        maxrow = BGS.Count
        maxcol = BGS.Item(0).Length

        '判断计算桩号范围是否合理
        If jszh < 0 Then
            MsgBox("桩号过小")
        ElseIf jszh > BGS.Item(maxrow - 1)(0) Then
            MsgBox("桩号过大")
        End If

        '寻找最近变坡点
        For i = 0 To maxrow
            If BGS.Item(i)(0) - jszh >= 0 Then
                If Math.Abs(BGS.Item(i)(0) - jszh) < Math.Abs(BGS.Item(i)(0) - jszh) Then
                    row = i - 1
                Else
                    row = i
                End If
                Exit For
            End If
        Next i

        '主要计算参数赋值
        zxzh = BGS.Item(row)(0)
        zxbg = BGS.Item(row)(1)
        r = BGS.Item(row)(2)

        '分最前、最后、以及中间部分三段计算所求桩号位置标高。
        If row = 0 Then
            hpd = (BGS.Item(row + 1)(1) - BGS.Item(row)(1)) / (BGS.Item(row + 1)(0) - BGS.Item(row)(0))
            CalBySQX = BGS.Item(row)(1) + hpd * (jszh - BGS.Item(row)(0))
        ElseIf row = maxrow - 1 Then
            qpd = (BGS.Item(row)(1) - BGS.Item(row - 1)(1)) / (BGS.Item(row)(0) - BGS.Item(row - 1)(0))
            CalBySQX = BGS.Item(row - 1)(1) + qpd * (jszh - BGS.Item(row - 1)(0))
        Else
            hpd = (BGS.Item(row + 1)(1) - BGS.Item(row)(1)) / (BGS.Item(row + 1)(0) - BGS.Item(row)(0))
            qpd = (BGS.Item(row)(1) - BGS.Item(row - 1)(1)) / (BGS.Item(row)(0) - BGS.Item(row - 1)(0))

            '凸曲线与凹曲线判断
            If qpd > hpd Then
                flag = 1
            Else
                flag = -1
            End If

            t = Math.Abs(hpd - qpd) * r / 2
            e = Math.Sqrt(r ^ 2 + t ^ 2) - r
            If jszh < zxzh - t Then
                CalBySQX = zxbg - (zxzh - jszh) * qpd
            ElseIf jszh > zxzh + t Then
                CalBySQX = zxbg + (jszh - zxzh) * hpd
            ElseIf jszh < zxzh Then
                CalBySQX = zxbg - qpd * (zxzh - jszh) - flag * (t - Math.Abs(zxzh - jszh)) ^ 2 / 2 / r
            Else
                CalBySQX = zxbg - hpd * (zxzh - jszh) - flag * (t - Math.Abs(zxzh - jszh)) ^ 2 / 2 / r
            End If
        End If

        CalBySQX = Format(CalBySQX, "0.000")
        'CalBySQX = Math.Round(CalBySQX, 3)
    End Function

    '点位置标高计算函数，输入参数（竖曲线list，点中心桩号，横坡数组（1.5,-1.5）左正右负为常规中高两边低路拱，点到中心线距离（左负右正））
    Public Function CalPointHeight(ByVal BGS As List(Of Array), ByVal jszh As Double, ByVal CrossSlope() As Double, pointdis As Double)
        If pointdis <= 0 Then
            CalPointHeight = CalBySQX(BGS, jszh) + CrossSlope(0) * pointdis * 0.01
        Else
            CalPointHeight = CalBySQX(BGS, jszh) + CrossSlope(1) * pointdis * 0.01
        End If
    End Function

End Module
