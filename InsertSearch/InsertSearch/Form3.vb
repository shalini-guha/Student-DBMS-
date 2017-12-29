Imports System.Data.OleDb
Public Class Form3
    Public Property x As Integer
    Public Property y As String

    Dim con As New OleDbConnection

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.ConnectionString = "provider=microsoft.ace.oledb.12.0;data source=C:\Users\Zlalini\Desktop\Database1.accdb"
        con.Open()
        shalini()
    End Sub
    Private Sub shalini()
        Dim dt As New DataTable
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("select * from Table1", con)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
    End Sub
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim index As Integer
        index = e.RowIndex
        Dim selectedRow As DataGridViewRow
        selectedRow = DataGridView1.Rows(index)
        TextBox4.Text = selectedRow.Cells(0).Value.ToString
        TextBox5.Text = selectedRow.Cells(1).Value.ToString
        TextBox1.Text = selectedRow.Cells(7).Value.ToString
        TextBox2.Text = selectedRow.Cells(8).Value.ToString
        TextBox3.Text = selectedRow.Cells(9).Value.ToString
        TextBox6.Text = selectedRow.Cells(10).Value.ToString
        TextBox7.Text = selectedRow.Cells(11).Value.ToString
        TextBox8.Text = selectedRow.Cells(12).Value.ToString
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim command As New OleDbCommand
        Dim t, pe As Integer
        TextBox5.Text = x
        TextBox4.Text = y
        command = New OleDbCommand("update Table1 set Roll=@roll,Name=@name,Physics=@physics,Chemistry=@chemistry,Maths=@maths,Total=@total,Percentage=@percentage,Grade=@grade where Roll=@roll", con)
        command.Parameters.Add("@roll", OleDbType.Integer).Value = TextBox5.Text
        command.Parameters.Add("@name", OleDbType.VarChar).Value = TextBox4.Text
        command.Parameters.Add("@physics", OleDbType.Integer).Value = TextBox1.Text
        command.Parameters.Add("@chemistry", OleDbType.Integer).Value = TextBox2.Text
        command.Parameters.Add("@maths", OleDbType.Integer).Value = TextBox3.Text
        Dim g As String
        Dim p, c, m As Integer
        p = (TextBox1.Text)
        c = (TextBox2.Text)
        m = (TextBox3.Text)
        t = p + c + m
        pe = (t / 300) * 100
        If pe >= 90 Then
            g = "O"
        ElseIf pe >= 80 And pe < 90 Then
            g = "A+"
        ElseIf pe >= 70 And pe < 80 Then
            g = "A"
        ElseIf pe >= 60 And pe < 70 Then
            g = "B"
        ElseIf pe >= 50 And pe < 60 Then
            g = "C"
        Else
            g = "D"

        End If
        TextBox6.Text = t
        TextBox7.Text = pe
        TextBox8.Text = g
        command.Parameters.Add("@total", OleDbType.Integer).Value = TextBox6.Text
        command.Parameters.Add("@percentage", OleDbType.Integer).Value = TextBox7.Text
        command.Parameters.Add("@grade", OleDbType.VarChar).Value = TextBox8.Text
        If command.ExecuteNonQuery = 1 Then
            MsgBox("Record Inserted")
        Else
            MsgBox("Error in inserting Data")
        End If
        shalini()
    End Sub


End Class