Imports IWshRuntimeLibrary
Imports System.IO
Public Class FormAdd
  Private Sub FormAddEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    GetProcesses()
    Button1.Enabled = False
    Label3.ForeColor = Color.Red
    TextBox2.ReadOnly = True
    Label3.Text = ""
    Button1.Text = "Add"
  End Sub

  Public Sub GetProcesses()
    Dim idProcess As Integer = 0
    Try
      ListBox1.Items.Clear()

      Dim p As Process
      Dim duplikat As Boolean = False
      For Each p In Process.GetProcesses
        If Not p.MainWindowTitle = "" Then
          If Not p.ProcessName.Contains("Windows") Then
            'If Not p.MainWindowTitle = "Task Manager" Then
            For Each l In ListBox2.Items()
              If p.ProcessName = l.ToString Then
                duplikat = True
              Else
                duplikat = False
              End If
            Next
            If duplikat = False Then
              ListBox1.Items.Add(p.MainWindowTitle)
              ListBox2.Items.Add(p.ProcessName)
              'End If
            End If
          End If
        End If
      Next
    Catch ex As Exception
      MsgBox(ex, MsgBoxStyle.Critical, DialogResult)
    End Try

    ListBox1.Items.Add("Custom...")
  End Sub

  Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
    Select Case ListBox1.SelectedItem
      Case "Custom..."
        MsgBox("ok")
      Case Else
        For Each p In Process.GetProcesses
          Try
            If p.MainWindowTitle = ListBox1.SelectedItem.ToString Then
              TextBox1.Text = p.ProcessName
              TextBox2.Text = p.Modules(0).FileName.ToString
              Label3.Text = ""
              Button1.Enabled = True
            End If
          Catch ex As Exception
            TextBox2.Text = ""
            If p.MainWindowTitle = ListBox1.SelectedItem.ToString Then
              Label3.Text = p.ProcessName & " is invalid path"
              Button1.Enabled = False
            End If
          End Try
        Next
    End Select
  End Sub

  Private Sub makeShortcut()
    Dim WshShell As WshShell = New WshShell()
    Dim ShortcutPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Startup)

    Dim Shortcut As IWshShortcut = CType(WshShell.CreateShortcut(Path.Combine(ShortcutPath, TextBox1.Text & "_startup") & ".lnk"), IWshShortcut)
    Shortcut.TargetPath = TextBox2.Text
    Shortcut.Arguments = "-m"
    Shortcut.WorkingDirectory = Path.GetDirectoryName(TextBox2.Text)
    Shortcut.Description = ""
    Shortcut.Save()
  End Sub

  Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    makeShortcut()
    Me.Close()
    Form1.Loadas()
  End Sub
End Class