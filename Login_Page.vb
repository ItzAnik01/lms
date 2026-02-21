Imports System.Data.OleDb

Public Class Login_Page

    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\.\.\..\College.mdb;")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            con.Open()

            Dim cmd As New OleDbCommand("SELECT Designation FROM Register WHERE UserName=? AND Password=?", con)
            cmd.Parameters.AddWithValue("?", TextBox1.Text)
            cmd.Parameters.AddWithValue("?", TextBox2.Text)

            Dim dr As OleDbDataReader = cmd.ExecuteReader()

            If dr.Read() Then
                MessageBox.Show("Login Successful")

                ' Create LMS instance and pass role and username
                Dim LMS As New LMS()
                LMS.userRole = dr("Designation").ToString()
                LMS.currentUser = TextBox1.Text

                ' Show LMS as modal
                Me.Hide()
                LMS.ShowDialog()

                ' After LMS closes, exit the app
                'Application.Exit()
            Else
                MessageBox.Show("Invalid Username or Password")
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox1.Focus()
            End If

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub


End Class
