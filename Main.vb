﻿Public Class Main
    Inherits ApplicationContext

    Public WithEvents Tray As NotifyIcon
    Private WithEvents MainMenu As ContextMenuStrip
    Private WithEvents mnuDisplayForm As ToolStripMenuItem
    Private WithEvents mnuSep1 As ToolStripSeparator
    Private WithEvents mnuExit As ToolStripMenuItem
    Private frm As New Form1

    Public Sub New()
        loading_start = Date.Now
        System.IO.Directory.SetCurrentDirectory(Application.StartupPath)
        ini.path = (".\ZametkeR.ini")

        'Initialize tray context menu
        mnuDisplayForm = New ToolStripMenuItem("Show")
        mnuSep1 = New ToolStripSeparator()
        mnuExit = New ToolStripMenuItem("Exit")
        MainMenu = New ContextMenuStrip
        MainMenu.Items.AddRange(New ToolStripItem() {mnuDisplayForm, mnuSep1, mnuExit})

        'Initialize the tray
        Tray = New NotifyIcon
        Tray.Icon = My.Resources.alpha
        Tray.ContextMenuStrip = MainMenu
        Tray.Text = "ZametkeR"

        Dim t = ini.IniReadValue("Main", "Tray_Use")
        If t = "0" Or t.ToUpper = "FALSE" Or t = "" Then
            frm.Show()
            Exit Sub
        End If

        tray_use = True
        Tray.Visible = True

        t = ini.IniReadValue("Main", "Tray_Start")
        If t = "0" Or t.ToUpper = "FALSE" Or t = "" Then frm.Show() Else tray_start = True
    End Sub

#Region " Event handlers "
    Private Sub AppContext_ThreadExit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ThreadExit
        'Guarantees that the icon will not linger.
        Tray.Visible = False
    End Sub

    Private Sub Tray_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Tray.DoubleClick
        If Not frm.Visible Then
            If frm.IsDisposed Then frm = New Form1
            frm.Show() : frm.Activate()
        End If
    End Sub

    Private Sub mnuDisplayForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDisplayForm.Click
        If Not frm.Visible Then
            If frm.IsDisposed Then frm = New Form1
            frm.Show() : frm.Activate()
        End If
    End Sub

    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Application.Exit()
    End Sub
#End Region

End Class

Public Class NodeSorter
    Implements IComparer(Of TreeNode)

    Public Function Compare(x As TreeNode, y As TreeNode) As Integer Implements IComparer(Of TreeNode).Compare
        'Return x.Text.CompareTo(y.Text)
        Return String.Compare(x.Text, y.Text)
    End Function
End Class