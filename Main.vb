Public Class Main
    Inherits ApplicationContext

    Public WithEvents Tray As NotifyIcon
    Private WithEvents MainMenu As ContextMenuStrip
    Private WithEvents mnuDisplayForm As ToolStripMenuItem
    Private WithEvents mnuSep1 As ToolStripSeparator
    Private WithEvents mnuExit As ToolStripMenuItem
    Private frm As New Form1

    Public Sub New()
        loading_start = Date.Now
        ini.path = (".\ZametkeR.ini")

        'Initialize the menus
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
