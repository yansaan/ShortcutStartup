'Imports IWshRuntimeLibrary
Imports System.IO
Imports System.Environment
Imports IWshRuntimeLibrary

Public Class Form1
  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Loadas()
    Button3.Enabled = False
  End Sub

  Public Shared Function GetLnkTarget(lnkPath As String) As String
    Dim shl = New Shell32.Shell()
    ' Move this to class scope
    lnkPath = System.IO.Path.GetFullPath(lnkPath)
    Dim dir = shl.[NameSpace](System.IO.Path.GetDirectoryName(lnkPath))
    Dim itm = dir.Items().Item(System.IO.Path.GetFileName(lnkPath))
    Dim lnk = DirectCast(itm.GetLink, Shell32.ShellLinkObject)
    Return lnk.Target.Path
  End Function

  Public Sub Loadas()
    ListBox1.Items.Clear()
    For Each f As String In Directory.GetFiles(GetFIle.dir, "*_startup.lnk")
      If IO.File.Exists(GetLnkTarget(f)) Then
        Dim files = Path.GetFileNameWithoutExtension(f).Replace("_startup", "")
        ListBox1.Items.Add(files)
      Else
        MsgBox("Program is not found, this program automaticly deleted!", vbCritical, "Error")
        IO.File.Delete(Path.Combine(dir, f) & ".lnk")
      End If
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
      IO.File.Delete(Path.Combine(dir, ListBox1.SelectedItem & "_startup") & ".lnk")
      Loadas()
      Button3.Enabled = False
    End If
  End Sub
End Class
