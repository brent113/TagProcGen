Public Class Main

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Path.Text = My.Settings.SavedPath
    End Sub

    Private Sub Main_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        Me.Path.Text = files(0)
        My.Settings.SavedPath = files(0)
    End Sub

    Private Sub Main_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Browse_Click(sender As Object, e As EventArgs) Handles Browse.Click
        Dim Dir As String = ""
        If Path.Text.Length Then Dir = IO.Path.GetDirectoryName(Path.Text)
        OpenFileDialog1.InitialDirectory = Dir

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Path.Text = OpenFileDialog1.FileName
            My.Settings.SavedPath = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Gen_Click(sender As Object, e As EventArgs) Handles Gen.Click
        Gen.Enabled = False

        If Not My.Computer.FileSystem.FileExists(Path.Text) Then
            MsgBox("File does not exist")
            Return
        End If

        Generate(Path.Text)

        Gen.Enabled = True
    End Sub
End Class