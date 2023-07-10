Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Textbox1.text = "C:\Users\"
    End Sub
  
    Public Function FileCount(ByVal DirectoryNAme As String) As Long
        Dim objFS As New Scripting.FileSystemObject
        Dim objFolder As Scripting.Folder
        If objFS.FolderExists(DirectoryNAme) Then
            objFolder = objFS.GetFolder(DirectoryNAme)
            FileCount = objFolder.Files.Count
        End If
   End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label1.Text = FileCount(Me.TextBox1.Text).ToString
    End Sub
End Class
