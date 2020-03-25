Option Explicit On
Public Class MainForm
    Const TB = "fZodR9XQDSUm21yCkr6zBqiveYah8bt4xsWpHnJE7jL5VG3guMTKNPAwcF"
    Const NumX = 177451812
    Const NumA = 100618342136696320

    'Dim BThread As Threading.Thread

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        'BThread = New Threading.Thread(AddressOf AV2BV)
    End Sub

    Public Function AV2BV(A As Long) As String
        Dim n， x As Long, i As Integer
        Dim t() As String, f(9) As String
        AV2BV = "0000000000"
        Try
            x = (A Xor NumX) + NumA '异或运算
            For i = 0 To 9
                n = x \ (58 ^ i) Mod 58
                f(i) = n
            Next
            t = SplitString(TB)
            For i = 0 To 9
                f(i) = Replace(f(i), f(i), t(Val(f(i)) + 1))
            Next
            '排序
            AV2BV = f(6) & f(2) & f(4) & f(8) & f(5) & f(9) & f(3) & f(7) & f(1) & f(0)
        Catch e As Exception
            MessageBox.Show(e.Message)
        End Try
    End Function

    Public Function D2B(Num As Long) As String '二进制转换(未使用)
        D2B = ""
        Do While Num > 0
            D2B = Num Mod 2 & D2B
            Num \= 2
        Loop
    End Function

    Public Function SplitString(Text As String) As String() '切割字符串
        Dim i As Integer, s(58) As String
        For i = 1 To 58
            s(i) = Mid(Text, i, 1)
        Next
        SplitString = s
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox1.Text = "" Then Exit Sub '不能为空
            If InStr(LCase(TextBox1.Text), "av", CompareMethod.Text) Then
                TextBox2.Text = "BV" & AV2BV(Strings.Right(TextBox1.Text, Len(TextBox1.Text) - 2))
            Else
                TextBox2.Text = "BV" & AV2BV((TextBox1.Text))
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        '输出文本框不能被用户编辑
        If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
            e.Handled = True
        End If
    End Sub

    Private Sub MainForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        End
    End Sub
End Class
