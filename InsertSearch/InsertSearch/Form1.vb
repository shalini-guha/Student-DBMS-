Public Class Form1
    Dim f As New Form2
    Dim g As New Form3
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a, b As String
        a = TextBox1.Text
        b = TextBox2.Text
        If a = "uem" And b = "Shalini" Then
            MsgBox("LOGGED IN SUCCESSFULLY", MsgBoxStyle.Information, "VALID LOG IN")
            Me.Hide()
            f.Show()
        Else
            If a = "" And b = "" Then
                MsgBox("NO USER NAME OR PASSWORD", MsgBoxStyle.Critical, "INVALID LOG IN")
            Else
                If a = "uem" And b = "" Then
                    MsgBox("NO PASSWORD", MsgBoxStyle.Critical, "NO PASSWORD")
                Else
                    If a = "" And b = "Shalini" Then
                        MsgBox("NO USER NAME", MsgBoxStyle.Critical, "NO USER NAME")
                    Else
                        MsgBox("FAKE USER", MsgBoxStyle.Critical, "INVALID USER")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
End Class