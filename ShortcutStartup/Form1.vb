'Imports IWshRuntimeLibrary
Imports System.IO
Imports System.Environment
Public Class Form1
  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Loadas()
    Button3.Enabled = False
  End Sub

  Public Sub Loadas()
    ListBox1.Items.Clear()
    For Each file As String In Directory.GetFiles(GetFIle.dir, "*_startup.lnk")
      Dim files = Path.GetFileNameWithoutExtension(file).Replace("_startup", "")
      ListBox1.Items.Add(files)
    Next
  End Sub



  Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
    If ListBox1.SelectedIndex = -1 Then
      Button3.Enabled = False
    Else
      Button3.Enabled = True
    End If
  End Sub

  Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    FormAdd.ShowDialog()
  End Sub

  Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    Dim del = MsgBox("Are you sure delete " & ListBox1.SelectedItem & " from startup?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Sure?")
    If del = MsgBoxResult.Yes Then
      File.Delete(Path.Combine(GetFolderPath(SpecialFolder.Startup), ListBox1.SelectedItem & "_startup") & ".lnk")
      Loadas()
      Button3.Enabled = False
    End If
  End Sub
End Class
