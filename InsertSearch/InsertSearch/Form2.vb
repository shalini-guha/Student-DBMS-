Imports System.Data.OleDb
Imports System.Net.Mail

Public Class Form2
    Dim smtp As New SmtpClient
    Dim message As New MailMessage
    Dim con As New OleDbConnection
    Dim mail As New MailMessage
    Dim bytes() As Byte
    
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.ConnectionString = "provider=microsoft.ace.oledb.12.0;Data Source=C:\Users\Zlalini\Desktop\Database1.accdb"
        con.Open()
        shalini()
    End Sub
    Private Sub shalini()
        Dim dt As New DataTable
        Dim da As New OleDbDataAdapter
        Dim command As New OleDbCommand
        Dim h As Integer
        command = New OleDbCommand("Select COUNT(*) FROM Table1", con)
        h = command.ExecuteScalar()
        command = New OleDbCommand("Select COUNT(*) FROM Table2", con)
        h = h + command.ExecuteScalar()
        command = New OleDbCommand("Select COUNT(*) FROM Table3", con)
        h = h + command.ExecuteScalar()
        If h >= 0 And h <= 2 Then
            da = New OleDbDataAdapter("select * from Table1", con)
            RadioButton6.Checked = True
        ElseIf h > 2 And h <= 4 Then
            da = New OleDbDataAdapter("select * from Table2", con)
            RadioButton4.Checked = True
        ElseIf h > 4 And h <= 6 Then
            da = New OleDbDataAdapter("select * from Table3", con)
            RadioButton5.Checked = True
        End If
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
    End Sub
    Private Sub column()
        Dim dt As New DataTable
        Dim da As New OleDbDataAdapter
        Dim command As New OleDbCommand
        If RadioButton6.Checked = True Then
            da = New OleDbDataAdapter("select * from Table1", con)
        ElseIf RadioButton4.Checked = True Then
            da = New OleDbDataAdapter("select * from Table2", con)
        ElseIf RadioButton5.Checked = True Then
            da = New OleDbDataAdapter("select * from Table3", con)
        End If

        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim opf As New OpenFileDialog
        If opf.ShowDialog = DialogResult.OK Then
            PictureBox1.Image = Image.FromFile(opf.FileName)
        End If
    End Sub



    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dt As New DataTable
        Dim da As New OleDbDataAdapter
        Dim command As OleDbCommand
        Dim h, roll As Integer
        Dim bytes() As Byte
        Dim shalini1() As Byte

        Dim ms As New System.IO.MemoryStream
        command = New OleDbCommand("Select COUNT(*) FROM Table1", con)
        h = command.ExecuteScalar()
        command = New OleDbCommand("Select COUNT(*) FROM Table2", con)
        h = h + command.ExecuteScalar()
        command = New OleDbCommand("Select COUNT(*) FROM Table3", con)
        h = h + command.ExecuteScalar()
        If h >= 0 And h <= 2 Then
            command = New OleDbCommand("delete from Table1 where Roll=@roll", con)
        ElseIf h > 2 And h <= 4 Then
            command = New OleDbCommand("delete from Table2 where Roll=@roll", con)
        ElseIf h > 4 And h <= 6 Then
            command = New OleDbCommand("delete from Table3 where Roll=@roll", con)
        End If

        command.Parameters.Add("@roll", OleDbType.Integer).Value = TextBox2.Text
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        DateTimePicker1.Text = ""
        PictureBox1.Image = Nothing
        da = New OleDbDataAdapter(command)
        da.Fill(dt)
        MsgBox("Data Deleted")
        shalini()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim dt As New DataTable
        Dim da As New OleDbDataAdapter
        Dim command As OleDbCommand
        Dim h, k As String
        Dim ms As New System.IO.MemoryStream
        Dim bytes() As Byte

        If RadioButton6.Checked = True Then
            command = New OleDbCommand("select * from Table1 where roll=@roll", con)
        ElseIf RadioButton5.Checked = True Then
            command = New OleDbCommand("select * from Table2 where roll=@roll", con)
        ElseIf RadioButton4.Checked = True Then
            command = New OleDbCommand("select * from Table3 where roll=@roll", con)
        End If
        command.Parameters.Add("@roll", OleDbType.Integer).Value = TextBox2.Text

        da = New OleDbDataAdapter(command)
        da.Fill(dt)
        If dt.Rows.Count > 0 Then
            h = dt.Rows(0)(2).ToString
            TextBox1.Text = dt.Rows(0)(0).ToString
            TextBox2.Text = dt.Rows(0)(1).ToString
            TextBox3.Text = dt.Rows(0)(5).ToString
            If h = "B.TECH" Then
                CheckBox1.Checked = True
                CheckBox2.Checked = False
                CheckBox3.Text = False
                CheckBox4.Checked = False
            ElseIf h = "M.TECH" Then
                CheckBox1.Checked = False
                CheckBox2.Checked = True
                CheckBox3.Text = False
                CheckBox4.Checked = False
            ElseIf h = "BCA" Then
                CheckBox1.Checked = False
                CheckBox2.Checked = False
                CheckBox3.Checked = True
                CheckBox4.Checked = False
            ElseIf h = "MCA" Then
                CheckBox1.Checked = False
                CheckBox2.Checked = False
                CheckBox3.Text = False
                CheckBox4.Checked = True
            End If

            k = dt.Rows(0)(3).ToString
            If k = "Male" Then
                RadioButton1.Checked = True
            ElseIf k = "Female" Then
                RadioButton2.Checked = True
            Else
                RadioButton3.Checked = True
            End If
            DateTimePicker1.Text = dt.Rows(0)(4).ToString

            If dt.Rows(0)(6) Is DBNull.Value Then
                PictureBox1.Image = Nothing
            Else
                bytes = dt.Rows(0)(6)
                PictureBox1.Image = Image.FromStream(New System.IO.MemoryStream(bytes))
            End If
        Else
            MsgBox("Roll Not Found")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim command As New OleDbCommand
        Dim h = 0, i As Integer
        Dim shalini1() As Byte
        Dim ms As New System.IO.MemoryStream
        command = New OleDbCommand("Select COUNT(*) FROM Table1", con)
        h = command.ExecuteScalar()
        command = New OleDbCommand("Select COUNT(*) FROM Table2", con)
        h = h + command.ExecuteScalar()
        command = New OleDbCommand("Select COUNT(*) FROM Table3", con)
        h = h + command.ExecuteScalar()
        If h >= 0 And h < 2 Then
            command = New OleDbCommand("insert into Table1(Roll,Name,Course,Gender,DateofBirth,Email,Picture)values(@roll,@stname,@course,@gender,@dtbirth,@email,@picture)", con)
            i = 0
        ElseIf h >= 2 And h < 4 Then
            command = New OleDbCommand("insert into Table2(Roll,Name,Course,Gender,DateofBirth,Email,Picture)values(@roll,@stname,@course,@gender,@dtbirth,@email,@picture)", con)
            i = 2
        ElseIf h >= 4 And h < 6 Then
            command = New OleDbCommand("insert into Table3(Roll,Name,Course,Gender,DateofBirth,Email,Picture)values(@roll,@stname,@course,@gender,@dtbirth,@email,@picture)", con)
            i = 4
        End If
        command.Parameters.Add("@roll", OleDbType.VarChar).Value = (h + 1) - i
        command.Parameters.Add("@stname", OleDbType.VarChar).Value = TextBox1.Text
        If CheckBox1.Checked = True Then
            command.Parameters.Add("@stream", OleDbType.VarChar).Value = "B.TECH"
        ElseIf CheckBox2.Checked = True Then

            command.Parameters.Add("@stream", OleDbType.VarChar).Value = "M.TECH"
        ElseIf CheckBox3.Checked = True Then

            command.Parameters.Add("@stream", OleDbType.VarChar).Value = "BCA"
        Else

            command.Parameters.Add("@stream", OleDbType.VarChar).Value = "MCA"
        End If
        If RadioButton1.Checked = True Then
            command.Parameters.Add("@gender", OleDbType.VarChar).Value = "Male"
        ElseIf RadioButton2.Checked = True Then

            command.Parameters.Add("@gender", OleDbType.VarChar).Value = "Female"
        Else

            command.Parameters.Add("@gender", OleDbType.VarChar).Value = "Others"

        End If

        DateTimePicker1.Text = DateTimePicker1.Value
        command.Parameters.Add("@dtbirth", OleDbType.VarChar).Value = DateTimePicker1.Value
        command.Parameters.Add("@email", OleDbType.VarChar).Value = TextBox3.Text
        If PictureBox1.Image Is Nothing Then
            command.Parameters.Add("@picture", OleDbType.Binary).Value = DBNull.Value
        Else
            PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)
            shalini1 = ms.GetBuffer
            command.Parameters.Add("@picture", OleDbType.Binary).Value = shalini1
            TextBox2.Text = (h + 1) - i
        End If
        If command.ExecuteNonQuery = 1 Then
            MsgBox("Record Inserted")
        Else
            MsgBox("Error in inserting Data")
        End If
        shalini()

    End Sub


    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim index As Integer
        Dim dt As New DataTable
        Dim shalini1() As Byte
        Dim ms As New System.IO.MemoryStream
        index = e.RowIndex
        Dim selectedRow As DataGridViewRow
        selectedRow = DataGridView1.Rows(index)
        TextBox2.Text = selectedRow.Cells(0).Value.ToString
        TextBox1.Text = selectedRow.Cells(1).Value.ToString
        If selectedRow.Cells(3).Value = "Male" Then
            RadioButton1.Checked = True
        Else
            RadioButton2.Checked = True
        End If
        If selectedRow.Cells(2).Value = "B.TECH" Then
            CheckBox1.Checked = True
            CheckBox2.Checked = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False
        ElseIf selectedRow.Cells(2).Value = "M.TECH" Then
            CheckBox2.Checked = True
            CheckBox1.Checked = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False
        ElseIf selectedRow.Cells(2).Value = "BCA" Then
            CheckBox2.Checked = False
            CheckBox1.Checked = False
            CheckBox3.Checked = True
            CheckBox4.Checked = False
        Else
            CheckBox4.Checked = True
            CheckBox2.Checked = False
            CheckBox1.Checked = False
            CheckBox3.Checked = False
        End If
        DateTimePicker1.Text = selectedRow.Cells(4).Value
        TextBox3.Text = selectedRow.Cells(5).Value

        If selectedRow.Cells(6).Value Is DBNull.Value Then
            PictureBox1.Image = Nothing
        Else
            shalini1 = selectedRow.Cells(6).Value
            PictureBox1.Image = Image.FromStream(New System.IO.MemoryStream(shalini1))
        End If
    End Sub



    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim command As New OleDbCommand
        Dim h, i As Integer
        Dim bytes() As Byte
        Dim shalini1() As Byte
        Dim ms As New System.IO.MemoryStream
        command = New OleDbCommand("Select COUNT(*) FROM Table1", con)
        h = command.ExecuteScalar()
        command = New OleDbCommand("Select COUNT(*) FROM Table2", con)
        h = h + command.ExecuteScalar()
        command = New OleDbCommand("Select COUNT(*) FROM Table3", con)
        h = h + command.ExecuteScalar()
        If h >= 0 And h <= 2 Then
            i = 0
            command = New OleDbCommand("update Table1 set Roll=@roll,Name=@stname,Course=@course,Gender=@gender,DateofBirth=@dtbirth,Email=@email,Picture=@picture where Roll=@roll", con)
        ElseIf h > 2 And h <= 4 Then
            i = 2
            command = New OleDbCommand("update Table2 set Roll=@roll,Name=@stname,Course=@course,Gender=@gender,DateofBirth=@dtbirth,Email=@email,Picture=@picture where Roll=@roll", con)
        ElseIf h > 4 And h <= 6 Then
            i = 4
            command = New OleDbCommand("update Table3 set Roll=@roll,Name=@stname,Course=@course,Gender=@gender,DateofBirth=@dtbirth,Email=@email,Picture=@picture where Roll=@roll", con)
        End If
        command.Parameters.Add("@roll", OleDbType.VarChar).Value = TextBox2.Text
        command.Parameters.Add("@stname", OleDbType.VarChar).Value = TextBox1.Text
        If CheckBox1.Checked = True Then
            command.Parameters.Add("@stream", OleDbType.VarChar).Value = "B.TECH"
        ElseIf CheckBox2.Checked = True Then

            command.Parameters.Add("@stream", OleDbType.VarChar).Value = "M.TECH"
        ElseIf CheckBox3.Checked = True Then

            command.Parameters.Add("@stream", OleDbType.VarChar).Value = "BCA"
        Else

            command.Parameters.Add("@stream", OleDbType.VarChar).Value = "MCA"
        End If
        If RadioButton1.Checked = True Then
            command.Parameters.Add("@gender", OleDbType.VarChar).Value = "Male"
        ElseIf RadioButton2.Checked = True Then

            command.Parameters.Add("@gender", OleDbType.VarChar).Value = "Female"
        Else

            command.Parameters.Add("@gender", OleDbType.VarChar).Value = "Others"

        End If

        DateTimePicker1.Text = DateTimePicker1.Value
        command.Parameters.Add("@dtbirth", OleDbType.VarChar).Value = DateTimePicker1.Value
        command.Parameters.Add("@email", OleDbType.VarChar).Value = TextBox3.Text
        If PictureBox1.Image Is Nothing Then
            command.Parameters.Add("@picture", OleDbType.Binary).Value = DBNull.Value
        Else


            PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)
            shalini1 = ms.GetBuffer
            command.Parameters.Add("@picture", OleDbType.Binary).Value = shalini1
        End If
        If command.ExecuteNonQuery = 1 Then
            MsgBox("Record Updated")
        Else
            MsgBox("Error in updating Data")
        End If
        column()


    End Sub
    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        Dim dt As New DataTable
        Dim da As New OleDbDataAdapter
        Dim command As New OleDbCommand
        da = New OleDbDataAdapter("select * from Table1", con)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        Dim dt As New DataTable
        Dim da As New OleDbDataAdapter
        Dim command As New OleDbCommand
        da = New OleDbDataAdapter("select * from Table2", con)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        Dim dt As New DataTable
        Dim da As New OleDbDataAdapter
        Dim command As New OleDbCommand
        da = New OleDbDataAdapter("select * from Table3", con)
        da.Fill(dt)
        DataGridView1.DataSource = dt.DefaultView
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim g As New Form3
        Me.Hide()
        g.Show()
        g.x = Val(TextBox2.Text)
        g.y = Val(TextBox1.Text)
    End Sub


End Class


