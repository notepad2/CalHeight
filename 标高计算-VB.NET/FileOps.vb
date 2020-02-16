Imports System.IO

'文件读取函数，将存储纵断面数据、地面线数据，横断面数据的文件读取到list(of array)对象里，返回list对象。
Module FileOps
    Public Function ReadFile(ByVal FileName As String)
        Dim sr As New StreamReader(FileName)
        Dim row As New List(Of Array)
        Dim data() As Double
        Dim line As String
        Dim str() As String
        Dim i, j, col As Integer

        line = sr.ReadLine()

        While (line <> Nothing)
            str = Split(line, " ")
            col = str.GetLength(0)
            ReDim data(col - 1)

            j = 0
            For j = 0 To 2
                data(j) = Val(str(j))
            Next j

            row.Add(data)
            line = sr.ReadLine()
            i += 1
        End While

        Return row
    End Function

End Module
