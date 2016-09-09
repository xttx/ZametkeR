Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.FileIO.FileSystem

Public Class Form1
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Dim refr As Boolean = False
    Dim order As New List(Of String)
    Dim crypt As New Cryptography("xus.ebr#usp//4u(hec*aT#hucharuzAC$cru")
    Dim reminders As New Reminders
    Dim content As New Dictionary(Of TreeNode, String)
    Dim content_bgcolor As New Dictionary(Of TreeNode, Color)
    Dim content_needCrypt As New List(Of TreeNode)
    Dim content_associatedNodes As New Dictionary(Of RichTextBox, TreeNode)
    Dim optionsTabPage As TabPage
    Dim lastChangedTab As TabPage = Nothing
    Dim lastOverwritenNode As TreeNode = Nothing
    Dim remindersText As String = ""
    Dim remindersAstralisText As String = ""
    Dim SAVE_PATH_START As String = ""
    Dim colorArray = {"Red", "Blue", "Green", "Yellow", "Orange", "White", "Black", "Aqua", "Magenta", "Custom ..."}
    Dim focusRtfNextTime As RichTextBox = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        refr = True
        optionsTabPage = TabControl1.TabPages(0)
        TabControl1.TabPages.Remove(optionsTabPage)
        For Each Font As FontFamily In System.Drawing.FontFamily.Families
            ToolStripComboBox2.Items.Add(Font.Name)
        Next

        'Load config
        UseTrayToolStripMenuItem.Checked = tray_use
        If UseTrayToolStripMenuItem.Checked Then
            StartInTrayToolStripMenuItem.Enabled = True
            MinimizeToTrayToolStripMenuItem.Enabled = True
            CloseToTrayToolStripMenuItem.Enabled = True
        Else
            StartInTrayToolStripMenuItem.Enabled = False
            MinimizeToTrayToolStripMenuItem.Enabled = False
            CloseToTrayToolStripMenuItem.Enabled = False
        End If
        StartInTrayToolStripMenuItem.Checked = tray_start
        Dim t = ini.IniReadValue("Main", "Tray_min")
        If t = "0" Or t.ToUpper = "FALSE" Or t = "" Then MinimizeToTrayToolStripMenuItem.Checked = False Else MinimizeToTrayToolStripMenuItem.Checked = True
        t = ini.IniReadValue("Main", "Tray_close")
        If t = "0" Or t.ToUpper = "FALSE" Or t = "" Then CloseToTrayToolStripMenuItem.Checked = False Else CloseToTrayToolStripMenuItem.Checked = True

        t = ini.IniReadValue("Main", "Autosave")
        If t = "0" Or t.ToUpper = "FALSE" Then AutosaveOnExitToolStripMenuItem.Checked = False
        t = ini.IniReadValue("Main", "Encrypt")
        If t <> "0" And t.ToUpper <> "FALSE" And t <> "" Then EncryptContentToolStripMenuItem.Checked = True
        t = ini.IniReadValue("Main", "Remember_last_pos_and_size")
        If t <> "0" And t.ToUpper <> "FALSE" And t <> "" Then RememberLastPositionAndSizeToolStripMenuItem.Checked = True
        If RememberLastPositionAndSizeToolStripMenuItem.Checked Then
            t = ini.IniReadValue("Main", "LastPos")
            If t <> "" Then
                Me.Left = CInt(t.Split(";"c)(0))
                Me.Top = CInt(t.Split(";"c)(1))
            End If
            t = ini.IniReadValue("Main", "LastSize")
            If t <> "" Then
                Me.Width = CInt(t.Split(";"c)(0))
                Me.Height = CInt(t.Split(";"c)(1))
            End If
        End If

        'options
        t = ini.IniReadValue("Main", "Dbl_click_on_tab")
        If t.ToUpper = "RENAME" Then RadioButton2.Checked = True Else RadioButton1.Checked = True

        t = ini.IniReadValue("Main", "Default_font_size")
        If t = "" Then ComboBox1.SelectedIndex = 2 Else ComboBox1.SelectedItem = t

        t = ini.IniReadValue("Main", "Ask_for_page_name")
        If t <> "0" And t.ToUpper <> "FALSE" And t <> "" Then CheckBox1.Checked = True

        t = ini.IniReadValue("Main", "Note_Order")
        If t.ToUpper = "KEEP" Or t = "" Then RadioButton3.Checked = True
        If t.ToUpper = "ALPHABETICAL" Then RadioButton4.Checked = True
        If t.ToUpper = "NOORDER" Then RadioButton5.Checked = True

        t = ini.IniReadValue("Main", "Use_Images")
        If t <> "0" And t.ToUpper <> "FALSE" And t <> "" Then CheckBox2.Checked = True

        t = ini.IniReadValue("Main", "Not_Load_All")
        If t <> "0" And t.ToUpper <> "FALSE" And t <> "" Then CheckBox3.Checked = True

        t = ini.IniReadValue("Main", "Delete_Confirm")
        If t = "0" Or t.ToUpper = "FALSE" Then CheckBox4.Checked = False

        t = ini.IniReadValue("Main", "Notes_Path").Trim
        If t <> "" Then TextBox1.Text = t : SAVE_PATH = t
        SAVE_PATH_START = SAVE_PATH

        t = ini.IniReadValue("Main", "Brose_in_options")
        If t.ToUpper = "LEFT" Then RadioButton6.Checked = True : CheckBox5.Enabled = True
        If t.ToUpper = "RIGHT" Then RadioButton7.Checked = True : CheckBox5.Enabled = True
        If t.ToUpper = "REPLACE" Then RadioButton8.Checked = True : CheckBox5.Enabled = False

        t = ini.IniReadValue("Main", "Brose_in_options_close")
        If t = "0" Or t.ToUpper = "FALSE" Then CheckBox5.Checked = False

        t = ini.IniReadValue("Main", "Brose_in_search")
        If t.ToUpper = "LEFT" Then RadioButton9.Checked = True : CheckBox6.Enabled = True
        If t.ToUpper = "RIGHT" Then RadioButton10.Checked = True : CheckBox6.Enabled = True
        If t.ToUpper = "REPLACE" Then RadioButton11.Checked = True : CheckBox6.Enabled = False

        t = ini.IniReadValue("Main", "Brose_in_search_close")
        If t <> "0" And t.ToUpper <> "FALSE" And t <> "" Then CheckBox6.Checked = True

        t = ini.IniReadValue("Main", "Use_Color_Presets")
        If t = "0" Or t.ToUpper = "FALSE" Then CheckBox7.Checked = False

        t = ini.IniReadValue("Main", "HotTrack")
        If t <> "0" And t.ToUpper <> "FALSE" And t <> "" Then CheckBox8.Checked = True

        t = ini.IniReadValue("Main", "Focus_RTF_on_tree_click")
        If t <> "0" And t.ToUpper <> "FALSE" And t <> "" Then CheckBox9.Checked = True

        refr = False
        If DirectoryExists(SAVE_PATH) Then load_recursively()
        If Not CheckBox2.Checked Then TreeView1.ImageList = Nothing

        'Add backups menu items
        If DirectoryExists(SAVE_PATH + "\!!!Backups") Then
            For Each f In GetDirectories(SAVE_PATH + "\!!!Backups")
                BackupsToolStripMenuItem.DropDownItems.Add(f.Substring(f.LastIndexOf("\") + 1))
            Next
        End If

        Dim d = DateTime.Now - loading_start
        'Label5.Text = "ZametkeR was loaded in " + d.Seconds.ToString + "." + d.Milliseconds.ToString + " seconds"
        Label5.Text = "ZametkeR was loaded in " + Math.Round(d.TotalSeconds, 3).ToString + " seconds"
    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If AutosaveOnExitToolStripMenuItem.Checked Then ToolStripButton1_Click(ToolStripButton1, New EventArgs)
        If refr Then Exit Sub
        If RememberLastPositionAndSizeToolStripMenuItem.Checked Then
            ini.IniWriteValue("Main", "LastPos", Me.Left.ToString + ";" + Me.Top.ToString)
            ini.IniWriteValue("Main", "LastSize", Me.Width.ToString + ";" + Me.Height.ToString)
        End If
        If Not UseTrayToolStripMenuItem.Checked Or Not CloseToTrayToolStripMenuItem.Checked Then Application.Exit()
    End Sub
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If WindowState = FormWindowState.Minimized Then
            If UseTrayToolStripMenuItem.Checked And MinimizeToTrayToolStripMenuItem.Checked Then
                refr = True
                Me.Close()
            End If
        End If
    End Sub

    'Load notes
    Private Sub load_recursively()
        Dim dir_root = GetDirectoryInfo(SAVE_PATH).FullName
        If Not dir_root.EndsWith("\") Then dir_root = dir_root + "\"

        Dim files = GetFiles(SAVE_PATH, FileIO.SearchOption.SearchAllSubDirectories, {"*.zam"}).ToList
        For i As Integer = files.Count - 1 To 0 Step -1
            If files(i).ToUpper.Contains("!!!BACKUPS") Then files.RemoveAt(i)
        Next
        If RadioButton3.Checked Or RadioButton4.Checked Then files.Sort()

        'ORDER
        If RadioButton3.Checked AndAlso FileExists(SAVE_PATH + "\!order.ord") Then
            order.Clear()

            Dim rdr As New StreamReader(SAVE_PATH + "\!order.ord")
            Do While Not rdr.EndOfStream
                order.Add(rdr.ReadLine)
            Loop
            rdr.Close()

            Dim tmp As New List(Of String)
            For Each o In order
                For Each f In files
                    If f.Substring(dir_root.Length).ToUpper = o.ToUpper + ".ZAM" Then
                        tmp.Add(f) : Exit For
                    End If
                Next
            Next
            If tmp.Count <> files.Count Then
                For Each f In files
                    If Not tmp.Contains(f) Then tmp.Add(f)
                Next
            End If
            files = tmp
        End If
        'ORDER END

        For Each f In files
            Dim node As TreeNode = Nothing
            Dim fname = f.Substring(f.LastIndexOf("\") + 1)
            Dim fnameNoExt = fname.Substring(0, fname.LastIndexOf("."))
            Dim dir = f.Substring(0, f.LastIndexOf("\"))
            dir = GetDirectoryInfo(dir).FullName
            If Not dir.EndsWith("\") Then dir = dir + "\"
            Dim dir_rel = dir.Substring(dir_root.Length).Trim
            If Not dir_rel.EndsWith("\") Then dir_rel = dir_rel + "\"

            If fnameNoExt = "Заметки ВК" Then
                f = f
            End If

            If dir_rel = "\" Then
                node = TreeView1.Nodes.Add(fnameNoExt, fnameNoExt) : content.Add(node, "") : node.ImageIndex = 3 : node.SelectedImageIndex = 3
            Else
                dir_rel += fnameNoExt
                For Each path As String In dir_rel.Split("\"c)
                    If node Is Nothing Then
                        node = TreeView1.Nodes(path)
                        If node Is Nothing Then node = TreeView1.Nodes.Add(path, path) : content.Add(node, "") : node.ImageIndex = 3 : node.SelectedImageIndex = 3
                    Else
                        If node.Nodes(path) Is Nothing Then
                            If node.ImageKey = "" Then node.ImageIndex = 2 : node.SelectedImageIndex = 2
                            node = node.Nodes.Add(path, path) : content.Add(node, "") : node.ImageIndex = 3 : node.SelectedImageIndex = 3
                        Else
                            node = node.Nodes(path)
                        End If
                    End If
                Next
            End If

            Dim isReminder As Boolean = node.Text.ToUpper = "REMINDERS"
            isReminder = isReminder Or (node.Parent IsNot Nothing AndAlso node.Parent.Text.ToUpper = "REMINDERS")
            If CheckBox3.Checked And Not isReminder Then
                content(node) = "{%%%UNLOADED%%%}"
            Else
                Dim tmp = load_content(f, node)
                content(node) = tmp

                'reminder check
                If node.Text.ToUpper = "REMINDERS" Then
                    node.BackColor = Color.LightBlue
                    Dim tmprtf As New RichTextBox
                    If tmp.ToUpper.StartsWith("{\RTF") Then
                        tmprtf.Rtf = tmp
                    Else
                        tmprtf.Text = tmp
                    End If
                    checkReminders(tmprtf)
                End If

                If node.Parent IsNot Nothing AndAlso node.Parent.Text.ToUpper = "REMINDERS" Then
                    node.ForeColor = Color.Aqua
                    Dim tmprtf As New RichTextBox
                    If tmp.ToUpper.StartsWith("{\RTF") Then
                        tmprtf.Rtf = tmp
                    Else
                        tmprtf.Text = tmp
                    End If
                    checkRemindersAstralis(tmprtf, node)
                End If
                'END reminder check
            End If

            'Need to be loaded unconditionally, because it show node colors
            load_bcg(f, node)

            'custom icon
            If FileExists(dir + fnameNoExt + ".png") Then
                Dim image = Bitmap.FromFile(dir + fnameNoExt + ".png")
                ImageList1.Images.Add(dir + fnameNoExt + ".png", image)
                node.ImageKey = dir + fnameNoExt + ".png"
                node.SelectedImageKey = dir + fnameNoExt + ".png"
            End If
        Next
        If TreeView1.Nodes.Count > 0 Then TreeView1.SelectedNode = TreeView1.Nodes(0)
    End Sub
    Private Function load_content(f As String, node As TreeNode) As String
        Dim w As New StreamReader(f)
        Dim tmp = w.ReadToEnd
        If tmp.StartsWith("CRYPF:") Then content_needCrypt.Add(node)
        If tmp.StartsWith("CRYPF:") Or tmp.StartsWith("CRYPR:") Then tmp = crypt.DecryptData(tmp.Substring(6))
        w.Close()
        Return tmp
    End Function
    Private Sub load_bcg(fZam As String, node As TreeNode)
        Dim bcg = fZam.Substring(0, fZam.Length - 4) + ".bcg"
        If FileExists(bcg) Then
            Dim w1 As New StreamReader(bcg)
            Dim cols = w1.ReadLine.Split("\"c)

            If cols.Count = 3 Then
                'new format
                Dim c = cols(0).Split(";"c)
                If CInt(c(0)) <> 0 Then node.ForeColor = Color.FromArgb(CInt(c(0)), CInt(c(1)), CInt(c(2)), CInt(c(3)))
                c = cols(1).Split(";"c)
                If CInt(c(0)) <> 0 Then node.BackColor = Color.FromArgb(CInt(c(0)), CInt(c(1)), CInt(c(2)), CInt(c(3)))
                c = cols(2).Split(";"c)
                Dim clr = Color.FromArgb(255, CInt(c(0)), CInt(c(1)), CInt(c(2)))
                If clr <> Color.White Then content_bgcolor.Add(node, clr)
            Else
                'old format
                Dim c = cols(0).Split(";"c)
                Dim clr = Color.FromArgb(255, CInt(c(0)), CInt(c(1)), CInt(c(2)))
                If clr <> Color.White Then content_bgcolor.Add(node, clr)
            End If

            w1.Close()
        End If
    End Sub
    Private Sub checkReminders(rtf As RichTextBox)
        reminders.parse(rtf.Text)

        If reminders.reminders.Count > 0 Then
            remindersText = ""
            For Each r In reminders.reminders
                remindersText += r.Key.ToString + " "
                If r.Value.status Then
                    remindersText += "Active"
                Else
                    remindersText += "Disabled"
                End If
                remindersText += vbCrLf
            Next
        End If
        If content_associatedNodes.ContainsKey(rtf) Then
            Dim sel_orig = rtf.SelectionStart
            Dim sel_orig_l = rtf.SelectionLength
            rtf.Select(0, rtf.Text.Length)
            rtf.SelectionColor = Color.Black
            For Each r In reminders.reminders
                Dim charPos = rtf.GetFirstCharIndexFromLine(r.Value.lineIndex)
                If charPos >= 0 Then
                    rtf.Select(charPos, r.Value.charcount)
                    refr = True : rtf.SelectionColor = Color.Turquoise : refr = False
                End If
            Next
            rtf.SelectionStart = sel_orig
            rtf.SelectionLength = sel_orig_l
        End If

        If remindersText.Trim <> "" Or remindersAstralisText.Trim <> "" Then
            ToolStripButton26.Visible = True
            ToolStripButton26.ToolTipText = remindersText.Trim + vbCrLf + vbCrLf + remindersAstralisText.Trim
        End If
    End Sub
    Private Sub checkRemindersAstralis(rtf As RichTextBox, node As TreeNode)
        reminders.parse_astralis(node, rtf.Text)

        If reminders.reminders_astralis.Count > 0 Then
            remindersAstralisText = ""
            For Each r In reminders.reminders_astralis
                remindersAstralisText += r.Value.dt.ToString + " "
                If r.Value.status Then
                    remindersAstralisText += "Active"
                Else
                    remindersAstralisText += "Disabled"
                End If
                remindersAstralisText += vbCrLf
            Next
        End If
        If content_associatedNodes.ContainsKey(rtf) Then
            Dim sel_orig = rtf.SelectionStart
            Dim sel_orig_l = rtf.SelectionLength
            rtf.Select(0, rtf.Text.Length)
            rtf.SelectionColor = Color.Black
            For Each r In reminders.reminders_astralis
                Dim charPos = rtf.GetFirstCharIndexFromLine(r.Value.lineIndex)
                If charPos >= 0 Then
                    rtf.Select(charPos, r.Value.charcount)
                    refr = True : rtf.SelectionColor = Color.Turquoise : refr = False
                End If
            Next
            rtf.SelectionStart = sel_orig
            rtf.SelectionLength = sel_orig_l
        End If

        If remindersText.Trim <> "" Or remindersAstralisText.Trim <> "" Then
            ToolStripButton26.Visible = True
            ToolStripButton26.ToolTipText = remindersText.Trim + vbCrLf + vbCrLf + remindersAstralisText.Trim
        End If
    End Sub

    'Hotkeys
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Alt And Not e.Control And e.KeyCode = Keys.A Then ToolStripButton2_Click(ToolStripButton2, New EventArgs)
        If e.Control And e.Alt And e.KeyCode = Keys.A Then ToolStripButton3_Click(ToolStripButton3, New EventArgs)
        If e.Control And e.Alt And e.KeyCode = Keys.D Then ToolStripButton4_Click(ToolStripButton4, New EventArgs)
        If e.Control And e.Alt And e.KeyCode = Keys.R Then ToolStripButton20_Click(ToolStripButton20, New EventArgs)

        If e.Control And e.KeyCode = Keys.S Then ToolStripButton1_Click(ToolStripButton1, New EventArgs)
        If e.Control And e.KeyCode = Keys.B Then ToolStripButton8_Click(ToolStripButton8, New EventArgs)
        If e.Control And e.KeyCode = Keys.I Then ToolStripButton9_Click(ToolStripButton9, New EventArgs)
        If e.Control And e.KeyCode = Keys.U Then ToolStripButton10_Click(ToolStripButton10, New EventArgs)
    End Sub

    'Treeview selected change - add or update tab page
    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        If refr Then Exit Sub
        openNote(e.Node)
    End Sub
    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        If TreeView1.SelectedNode Is Nothing Then Exit Sub
        If e.Button = MouseButtons.Left Then
            If e.Node IsNot TreeView1.SelectedNode AndAlso CheckBox9.Checked Then focusRtfNextTime = New RichTextBox
            If e.Node Is TreeView1.SelectedNode Then TreeView1_AfterSelect(TreeView1, New TreeViewEventArgs(e.Node))
        ElseIf e.Button = MouseButtons.Right Then
            If Clipboard.GetText.Trim = "" Then
                PastAsNewNoteToolStripMenuItem.Enabled = False
            Else
                PastAsNewNoteToolStripMenuItem.Enabled = True
            End If

            nodeForContextMenu = e.Node
            colorForContextMenu = e.Node.BackColor
            e.Node.BackColor = Color.LightGray
            ContextMenuStrip1.Show(Cursor.Position)

            'If e.Node Is TreeView1.SelectedNode Then Exit Sub
            'TreeView1.SelectedNode = e.Node
        End If
    End Sub
    Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
        If DirectCast(optionsTabPage.Controls("RadioButton1"), RadioButton).Checked Then
            'open new tab
            'openNewTabPage(e.Node)
            If lastChangedTab Is Nothing Then
                openNote(e.Node)
            Else
                'revert
                openNote(lastOverwritenNode)
                openNote(e.Node, True)
            End If
        Else
            'rename
            ToolStripButton20_Click(ToolStripButton20, New EventArgs)
        End If
    End Sub
    'Hotkeys on treeview
    Private Sub TreeView1_KeyDown(sender As Object, e As KeyEventArgs) Handles TreeView1.KeyDown
        If e.Control And e.KeyCode = Keys.V Then
            pastAsNewNote()
        End If
        If e.KeyCode = Keys.Delete And TreeView1.SelectedNode IsNot Nothing Then
            If CheckBox4.Checked Then
                If MsgBox("Delete " + TreeView1.SelectedNode.Text + " and all its children?", MsgBoxStyle.YesNo) <> MsgBoxResult.Yes Then Exit Sub
            End If
            ToolStripButton4_Click(ToolStripButton4, New EventArgs)
            TreeView1.Select()
        End If
        If e.KeyCode = Keys.F2 And TreeView1.SelectedNode IsNot Nothing Then
            ToolStripButton20_Click(ToolStripButton20, New EventArgs)
        End If
    End Sub
    'Tree got focus
    Sub TreeView1_Enter() Handles TreeView1.Enter
        If focusRtfNextTime IsNot Nothing Then focusRtfNextTime.Select() : focusRtfNextTime = Nothing
    End Sub
    'Hot track
    Private Sub TreeView1_MouseEnter(sender As Object, e As EventArgs) Handles TreeView1.MouseEnter
        If CheckBox8.Checked And Not TreeView1.Focused Then TreeView1.Select()
    End Sub

    'Open Note
    Private Sub openNote(node As TreeNode, Optional forceNewTab As Boolean = False)
        refr = True
        Dim rtf As RichTextBox
        Dim tabpage As TabPage
        Dim firstNoServiceTab As Integer = -1
        Dim lastNoServiceTab As Integer = -1

        'set first and last noService tabs and Check if this note already opened
        For Each tabpage In TabControl1.TabPages
            'If tabpage.Text = node.Text Then TabControl1.SelectedTab = tabpage : refr = False : Exit Sub
            If tabpage IsNot optionsTabPage And tabpage.Text.ToUpper <> "SEARCH" Then
                'Check if this note already opened
                rtf = tabpage.Controls.OfType(Of RichTextBox).First
                If content_associatedNodes.ContainsKey(rtf) AndAlso content_associatedNodes(rtf) Is node Then
                    TabControl1.SelectedTab = tabpage
                    If focusRtfNextTime IsNot Nothing Then rtf.Select() : focusRtfNextTime = rtf
                    refr = False : Exit Sub
                End If

                If firstNoServiceTab = -1 Then firstNoServiceTab = TabControl1.TabPages.IndexOf(tabpage)
                lastNoServiceTab = TabControl1.TabPages.IndexOf(tabpage)
            End If
        Next

        If forceNewTab Or firstNoServiceTab = -1 Then
            'if new tab needed
            lastChangedTab = Nothing
            lastOverwritenNode = Nothing

            Dim ind = TabControl1.TabPages.IndexOf(optionsTabPage)
            If ind < 0 Then
                TabControl1.TabPages.Add(node.Text)
                tabpage = TabControl1.TabPages(TabControl1.TabPages.Count - 1)
            Else
                TabControl1.TabPages.Insert(ind, node.Text)
                tabpage = TabControl1.TabPages(ind)
            End If
            TabControl1.SelectedIndex = TabControl1.TabPages.IndexOf(tabpage)

            rtf = openNote_insert(tabpage, node)
            'END if new tab needed
        Else
            tabpage = TabControl1.SelectedTab
            If tabpage Is optionsTabPage Then
                If (RadioButton6.Checked Or RadioButton7.Checked) And CheckBox5.Checked Then TabControl1.TabPages.Remove(tabpage)
                If RadioButton6.Checked Then tabpage = TabControl1.TabPages(firstNoServiceTab)
                If RadioButton7.Checked Then tabpage = TabControl1.TabPages(lastNoServiceTab)
                If RadioButton8.Checked Then
                    Dim ind As Integer = TabControl1.TabPages.IndexOf(tabpage)
                    TabControl1.TabPages.Insert(ind, node.Text)
                    TabControl1.TabPages.Remove(tabpage)
                    tabpage = TabControl1.TabPages(ind)
                    openNote_insert(tabpage, node)
                    TabControl1.SelectedTab = tabpage
                End If
            End If
            If tabpage.Text.ToUpper = "SEARCH" Then
                If (RadioButton9.Checked Or RadioButton10.Checked) And CheckBox6.Checked Then TabControl1.TabPages.Remove(tabpage)
                If RadioButton9.Checked Then tabpage = TabControl1.TabPages(firstNoServiceTab)
                If RadioButton10.Checked Then tabpage = TabControl1.TabPages(lastNoServiceTab)
                If RadioButton11.Checked Then
                    Dim ind As Integer = TabControl1.TabPages.IndexOf(tabpage)
                    TabControl1.TabPages.Insert(ind, node.Text)
                    TabControl1.TabPages.Remove(tabpage)
                    tabpage = TabControl1.TabPages(ind)
                    openNote_insert(tabpage, node)
                    TabControl1.SelectedTab = tabpage
                End If
            End If

            rtf = tabpage.Controls.OfType(Of RichTextBox).First

            lastChangedTab = tabpage
            lastOverwritenNode = content_associatedNodes(rtf)

            tabpage.Text = node.Text
            content_associatedNodes(rtf) = node
        End If

        If content(node).Contains("{%%%UNLOADED%%%") Then
            Dim f = SAVE_PATH + "\" + node.FullPath + ".zam"
            If FileExists(f) Then content(node) = load_content(f, node)
        End If

        'If content is empty, set default font size
        If content(node).Trim = "" Or Not content(node).ToUpper.StartsWith("{\RTF") Then
            Dim f = rtf.Font
            rtf.Font = New Font(f.FontFamily, CSng(ComboBox1.SelectedItem.ToString), f.Style)
            rtf.SelectionFont = New Font(f.FontFamily, CSng(ComboBox1.SelectedItem.ToString), f.Style)
            'Update toolbar controls. MAYBE THIS NEED TO BE MOVED OUT OF IF TO UPDATE CONTROLS UNCONDITIONALLY
            rtf_SelectionChanged(rtf, New EventArgs)
        End If

        If content(node).ToUpper.StartsWith("{\RTF") Then
            rtf.Rtf = content(node)
        Else
            rtf.Text = content(node)
        End If

        If content_bgcolor.ContainsKey(node) Then rtf.BackColor = content_bgcolor(node) Else rtf.BackColor = Color.White

        'Check reminder
        If node.Text.ToUpper = "REMINDERS" Then checkReminders(rtf)
        If node.Parent IsNot Nothing AndAlso node.Parent.Text.ToUpper = "REMINDERS" Then checkRemindersAstralis(rtf, node)

        'handle encrypt button
        TabControl1_SelectedIndexChanged(TabControl1, New EventArgs)

        'rtf to select
        If focusRtfNextTime IsNot Nothing Then rtf.Select() : focusRtfNextTime = rtf
        refr = False
    End Sub
    Private Function openNote_insert(tabpage As TabPage, node As TreeNode) As RichTextBox
        Dim rtf = New RichTextBox
        AddHandler rtf.TextChanged, AddressOf rtf_TextChanged
        AddHandler rtf.SelectionChanged, AddressOf rtf_SelectionChanged
        AddHandler rtf.MouseMove, AddressOf rtf_mouseMove
        rtf.Dock = DockStyle.Fill
        rtf.HideSelection = False
        rtf.AcceptsTab = True
        tabpage.Controls.Add(rtf)
        content_associatedNodes.Add(rtf, node)
        Return rtf
    End Function

    'Past as new
    Private Sub pastAsNewNote()
        Dim str = Clipboard.GetText.Trim
        Dim words = str.Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
        Dim newname As String = "New Page"
        If words.Count > 0 Then newname = words(0)
        If words.Count > 1 Then newname = words(0) + " " + words(1)

        'add page
        ToolStripButton2_Click(ToolStripButton2, New EventArgs, newname, Clipboard.GetText.Trim)
    End Sub

    'Change rtf text
    Private Sub rtf_TextChanged(sender As Object, e As EventArgs)
        If refr Then Exit Sub
        Dim rtf = DirectCast(sender, RichTextBox)
        Dim node = content_associatedNodes(rtf)
        content(node) = rtf.Rtf

        'reminder check - My format
        If node.Text.ToUpper = "REMINDERS" Then
            checkReminders(rtf)
        End If

        'reminder check - Astralis format
        If node.Parent IsNot Nothing AndAlso node.Parent.Text.ToUpper = "REMINDERS" Then
            checkRemindersAstralis(rtf, node)
        End If
    End Sub
    'Handle rtf image resize handlers and hotTrack
    Private Sub rtf_mouseMove(sender As Object, e As MouseEventArgs)
        Dim rtf = DirectCast(sender, RichTextBox)
        If rtf.SelectionType = RichTextBoxSelectionTypes.Object AndAlso rtf.SelectionLength = 1 Then
            Dim bounds As Rectangle
            bounds.Location = rtf.GetPositionFromCharIndex(rtf.SelectionStart)

            Dim lineIndex As Integer = rtf.GetLineFromCharIndex(rtf.SelectionStart)
            If rtf.Lines.Count < lineIndex + 1 Then Exit Sub
            bounds.Height = (rtf.GetPositionFromCharIndex(rtf.GetFirstCharIndexFromLine(lineIndex + 1))).Y - bounds.Y - 2
            bounds.Width = (rtf.GetPositionFromCharIndex(rtf.SelectionStart + rtf.SelectionLength)).X - bounds.X

            Dim centY = CInt((bounds.Height / 2) + bounds.Y)
            Dim centX = CInt((bounds.Width / 2) + bounds.X)

            Dim l, r, t, b As Boolean
            l = (e.X >= bounds.X And e.X < bounds.X + 7)
            r = (e.X > (bounds.X + bounds.Width) - 7 And e.X <= (bounds.X + bounds.Width))
            t = (e.Y >= bounds.Y And e.Y < bounds.Y + 7)
            b = (e.Y > (bounds.Y + bounds.Height) - 7 And e.Y <= (bounds.Y + bounds.Height))

            If (t And l) Or (b And r) Then
                rtf.Cursor = Cursors.SizeNWSE
            ElseIf (t And r) Or (b And l) Then
                rtf.Cursor = Cursors.SizeNESW
            ElseIf (t Or b) And (e.X > centX - 4 And e.X < centX + 4) Then
                rtf.Cursor = Cursors.SizeNS
            ElseIf (l Or r) And (e.y > centy - 4 And e.y < centy + 4) Then
                rtf.Cursor = Cursors.SizeWE
            Else
                rtf.Cursor = Cursors.IBeam
            End If
        End If
        If CheckBox8.Checked And Not rtf.Focused Then rtf.Select()
    End Sub

    'Draw tab close button
    Dim mouseHoverDrawn As Boolean = False
    Private Sub TabControl1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl1.DrawItem
        'Dim a = SendMessage(TabControl1.TabPages(e.Index).Handle, &H1300 + 49, IntPtr.Zero, 300)
        If e.Index > TabControl1.TabCount - 1 Then Exit Sub
        Dim text = TabControl1.TabPages(e.Index).Text
        If e.Graphics.MeasureString(text, e.Font).Width > 100 Then
            Do While e.Graphics.MeasureString(text, e.Font).Width > 90
                text = text.Substring(0, text.Length - 1)
            Loop
            text = text + "..."
        End If
        e.Graphics.FillRectangle(New SolidBrush(SystemColors.Control), e.Bounds)
        e.Graphics.DrawString("x", e.Font, Brushes.Black, e.Bounds.Right - 15, e.Bounds.Top + 4)
        e.Graphics.DrawString(text, e.Font, Brushes.Black, e.Bounds.Left + 12, e.Bounds.Top + 4)
        e.DrawFocusRectangle()
    End Sub
    'Draw tab close button - click
    Private Sub TabControl1_MouseDown(sender As Object, e As MouseEventArgs) Handles TabControl1.MouseDown
        For i As Integer = 0 To TabControl1.TabPages.Count - 1
            Dim r As Rectangle = TabControl1.GetTabRect(i)
            'Getting the position of the "x" mark.
            Dim closeButton As Rectangle = New Rectangle(r.Right - 15, r.Top + 4, 11, 10)
            If TabControl1.SelectedIndex <> i Then closeButton.X = closeButton.X - 2 : closeButton.Y = closeButton.Y + 2

            If closeButton.Contains(e.Location) Then
                If Not TabControl1.TabPages(i).Text = "Options" Then
                    Dim rtf = TabControl1.TabPages(i).Controls.OfType(Of RichTextBox).First
                    content_associatedNodes.Remove(rtf)
                End If
                TabControl1.TabPages.RemoveAt(i)
                Exit For
            End If
        Next
    End Sub
    'Draw tab close button - hover
    Private Sub TabControl1_MouseMove(sender As Object, e As MouseEventArgs) Handles TabControl1.MouseMove
        For i As Integer = 0 To TabControl1.TabPages.Count - 1
            Dim r As Rectangle = TabControl1.GetTabRect(i)
            Dim closeButton As Rectangle = New Rectangle(r.Right - 15, r.Top + 4, 11, 10)
            If TabControl1.SelectedIndex <> i Then closeButton.X = closeButton.X - 2 : closeButton.Y = closeButton.Y + 2
            If closeButton.Contains(e.Location) Then
                Dim g = TabControl1.CreateGraphics
                g.DrawRectangle(Pens.Black, closeButton)
                mouseHoverDrawn = True
            ElseIf mouseHoverDrawn Then
                Dim g = TabControl1.CreateGraphics
                g.DrawRectangle(New Pen(SystemColors.Control), closeButton)
            End If
        Next
    End Sub
    'Handle encrypt button
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        refr = True
        Dim ind = TabControl1.SelectedIndex
        If ind < 0 Then
            ToolStripButton25.Checked = False
            ToolStripButton25.Enabled = False
            ToolStripButton25.Image = ImageList1.Images(0)
        Else
            If TabControl1.TabPages(ind).Text = "Search" Then
                ToolStripButton25.Checked = False
                ToolStripButton25.Enabled = False
                ToolStripButton25.Image = ImageList1.Images(0)
            ElseIf TabControl1.TabPages(ind).Controls.Count = 1 Then
                ToolStripButton25.Enabled = True
                Dim rtf = TabControl1.TabPages(ind).Controls.OfType(Of RichTextBox).First
                Dim node = content_associatedNodes(rtf)
                If content_needCrypt.Contains(node) Then
                    ToolStripButton25.Checked = True
                    ToolStripButton25.Image = ImageList1.Images(1)
                Else
                    ToolStripButton25.Checked = False
                    ToolStripButton25.Image = ImageList1.Images(0)
                End If
            End If
        End If
        refr = False
    End Sub


    'Add page
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs, Optional new_name As String = "", Optional _content As String = "") Handles ToolStripButton2.Click
        Dim c As Integer = 0
        Dim node_main As TreeNode = nodeForContextMenu2
        If node_main Is Nothing Then node_main = TreeView1.SelectedNode

        Dim node As TreeNode = Nothing
        Dim default_text = "New Page"
        If new_name <> "" Then default_text = new_name

        If node_main Is Nothing Then
            default_text = getNewName(node_main, default_text)
            node = TreeView1.Nodes.Add(default_text, default_text)
        Else
            Dim tmp = node_main.Parent
            If tmp Is Nothing Then
                default_text = getNewName(node_main, default_text)
                node = TreeView1.Nodes.Add(default_text, default_text)
            Else
                default_text = getNewName(node_main, default_text)
                node = tmp.Nodes.Add(default_text, default_text)
            End If
        End If
        content.Add(node, _content)
        TreeView1.SelectedNode = node

        'rename if needed
        If CheckBox1.Checked Then
            If ToolStripButton20_Click(ToolStripButton20, New EventArgs) = "" Then
                'delete if renaming canceled
                ToolStripButton4_Click(ToolStripButton4, New EventArgs)
            End If
        End If
    End Sub
    'Add page hierarchy
    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Dim node_main As TreeNode = nodeForContextMenu2
        If node_main Is Nothing Then node_main = TreeView1.SelectedNode

        If node_main Is Nothing Then MsgBox("You must select a page to add child page") : Exit Sub
        If node_main.ImageKey = "" Then node_main.ImageIndex = 2 : node_main.SelectedImageIndex = 2

        Dim c As Integer = 0
        Dim default_text = "New Page"
        default_text = getNewName(node_main, default_text, True)
        Dim node = node_main.Nodes.Add(default_text, default_text)

        node_main.Expand()
        content.Add(node, "")
        TreeView1.SelectedNode = node

        'rename if needed
        If CheckBox1.Checked Then
            If ToolStripButton20_Click(ToolStripButton20, New EventArgs) = "" Then
                'delete if renaming canceled
                ToolStripButton4_Click(ToolStripButton4, New EventArgs)
            End If
        End If
    End Sub
    'Remove page
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim node_main As TreeNode = nodeForContextMenu2
        If node_main Is Nothing Then node_main = TreeView1.SelectedNode

        If node_main Is Nothing Then MsgBox("You must select a page to remove") : Exit Sub
        If node_main.Nodes.Count > 0 Then
            If MsgBox("Selected page has children. They will all be deleted. Are you sure?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Sub
        End If

        Dim ind = node_main.Index
        Dim parent = node_main.Parent
        removeRecursively(node_main)
        Dim t = node_main.Parent
        node_main.Remove()
        If t IsNot Nothing AndAlso t.Nodes.Count = 0 Then t.ImageIndex = 3 : t.SelectedImageIndex = 3

        'handle encrypt button
        TabControl1_SelectedIndexChanged(TabControl1, New EventArgs)

        'set selected node
        If TreeView1.Nodes.Count > 0 Then
            Dim node_sel As TreeNode = Nothing
            Dim nodes_p As TreeNodeCollection = Nothing
            If parent Is Nothing Then nodes_p = TreeView1.Nodes Else nodes_p = parent.Nodes
            If nodes_p.Count = 0 Then
                If parent IsNot Nothing Then node_sel = parent
            Else
                If nodes_p.Count > ind Then node_sel = nodes_p(ind) Else node_sel = nodes_p(nodes_p.Count - 1)
            End If
            If node_sel IsNot Nothing Then TreeView1.SelectedNode = node_sel
        End If
    End Sub
    Private Sub removeRecursively(node As TreeNode)
        For Each sub_node In node.Nodes
            removeRecursively(sub_node)
        Next

        Dim path = SAVE_PATH + "\" + node.FullPath
        If FileExists(path + ".zam") Then DeleteFile(path + ".zam")
        If DirectoryExists(path) Then DeleteDirectory(path, FileIO.DeleteDirectoryOption.DeleteAllContents)

        Dim tabPagesToRemove As New List(Of TabPage)
        For Each tabPage As TabPage In TabControl1.TabPages
            If tabPage Is optionsTabPage Then Continue For
            Dim rtf = tabPage.Controls.OfType(Of RichTextBox).First
            If content_associatedNodes(rtf) Is node Then
                tabPagesToRemove.Add(tabPage)
                content_associatedNodes.Remove(rtf)
            End If
        Next
        For Each tabPage In tabPagesToRemove
            TabControl1.TabPages.Remove(tabPage)
        Next

        content.Remove(node)
    End Sub
    'Rename page
    Private Function ToolStripButton20_Click(sender As Object, e As EventArgs) As String Handles ToolStripButton20.Click
        Return ToolStripButton20_rename_sub()
    End Function
    Private Function ToolStripButton20_rename_sub(Optional newname As String = "") As String
        Dim node_main As TreeNode = nodeForContextMenu2
        If node_main Is Nothing Then node_main = TreeView1.SelectedNode

        If node_main Is Nothing Then MsgBox("You must select a page to rename") : Return ""
        Dim node = node_main

        Dim nameIsEntered As Boolean = False

        If newname = "" Then newname = InputBox("New name", "Renaming page", node.Text).Trim : nameIsEntered = True
        Do While newname.Contains("/") Or newname.Contains("\") Or newname.Contains(":") Or newname.Contains("*") Or newname.Contains("?") Or newname.Contains("""") Or newname.Contains("<") Or newname.Contains(">") Or newname.Contains("|")
            newname = InputBox("Name contains illegal character (/, \, :, *, ?, """", >, <, |). Choose another name.", "Renaming page", node.Text).Trim : nameIsEntered = True
        Loop
        If newname = "" Then Return ""

        If newname.ToUpper <> node.Text.ToUpper Then
            newname = getNewName(node, newname)

            Dim old_path = SAVE_PATH + "\" + node.FullPath
            node.Name = newname
            node.Text = newname
            If Not Path.GetFullPath(old_path).ToUpper = Path.GetFullPath(SAVE_PATH + "\" + node.FullPath).ToUpper Then
                If DirectoryExists(Path.GetDirectoryName(old_path)) Then
                    For Each file In GetFiles(Path.GetDirectoryName(old_path), FileIO.SearchOption.SearchTopLevelOnly, {Path.GetFileName(old_path) + ".*"})
                        MoveFile(file, SAVE_PATH + "\" + node.FullPath + Path.GetExtension(file))
                    Next
                End If
            End If
            If DirectoryExists(old_path) Then RenameDirectory(old_path, node.Text)
        End If

        For Each tabPage As TabPage In TabControl1.TabPages
            If tabPage.Text = "Options" Then Continue For
            Dim rtf = tabPage.Controls.OfType(Of RichTextBox).First
            If content_associatedNodes(rtf) Is node Then
                tabPage.Text = newname
            End If
        Next

        'reminder check
        If node.Text.ToUpper.StartsWith("REMINDER") Then node.BackColor = Color.LightBlue

        'set focus
        If nameIsEntered Then
            If TabControl1.SelectedTab IsNot Nothing Then
                If TabControl1.SelectedTab.Text <> "Options" Then
                    Dim rtf = TabControl1.SelectedTab.Controls.OfType(Of RichTextBox).First
                    rtf.Select()
                End If
            End If
        End If
        Return "OK"
    End Function
    'Duplicate page
    Private Sub ToolStripButton21_Click(sender As Object, e As EventArgs) Handles ToolStripButton21.Click
        Dim node_main As TreeNode = nodeForContextMenu2
        If node_main Is Nothing Then node_main = TreeView1.SelectedNode

        If node_main Is Nothing Then MsgBox("You must select a page to clone") : Exit Sub

        Dim c As Integer = 0
        Dim node As TreeNode = node_main
        Dim default_text = node.Text
        If default_text.Contains(" (copy") Then default_text = default_text.Substring(0, default_text.IndexOf(" (copy"))
        Dim default_textE = default_text + " (copy)"

        Dim tmp = node.Parent
        If tmp Is Nothing Then
            Do While TreeView1.Nodes(default_textE) IsNot Nothing
                c += 1
                default_textE = default_text + " (copy " + c.ToString + ")"
            Loop
            node = TreeView1.Nodes.Add(default_textE, default_textE)
        Else
            Do While tmp.Nodes(default_textE) IsNot Nothing
                c += 1
                default_textE = default_text + " (copy " + c.ToString + ")"
            Loop
            node = tmp.Nodes.Add(default_textE, default_textE)
        End If
        content.Add(node, content(node_main))
        'TreeView1.SelectedNode = node_main
    End Sub
    'Move Up
    Private Sub ToolStripButton23_Click(sender As Object, e As EventArgs) Handles ToolStripButton23.Click
        If TreeView1.SelectedNode Is Nothing Then MsgBox("You must select a page to move") : Exit Sub
        Dim node = TreeView1.SelectedNode
        Dim nodes As TreeNodeCollection
        Dim ind = node.Index
        If ind = 0 Then Exit Sub
        refr = True
        If node.Parent Is Nothing Then nodes = TreeView1.Nodes Else nodes = node.Parent.Nodes
        nodes.RemoveAt(ind)
        nodes.Insert(ind - 1, node)
        refr = False
        TreeView1.SelectedNode = node
    End Sub
    'Move Down
    Private Sub ToolStripButton24_Click(sender As Object, e As EventArgs) Handles ToolStripButton24.Click
        If TreeView1.SelectedNode Is Nothing Then MsgBox("You must select a page to move") : Exit Sub
        Dim node = TreeView1.SelectedNode
        Dim nodes As TreeNodeCollection
        Dim ind = node.Index
        If node.Parent Is Nothing Then nodes = TreeView1.Nodes Else nodes = node.Parent.Nodes
        If ind = nodes.Count - 1 Then Exit Sub
        refr = True
        nodes.RemoveAt(ind)
        nodes.Insert(ind + 1, node)
        refr = False
        TreeView1.SelectedNode = node
    End Sub
    'Move Left
    Private Sub ToolStripButton28_Click(sender As Object, e As EventArgs) Handles ToolStripButton28.Click
        If TreeView1.SelectedNode Is Nothing Then MsgBox("You must select a page to move") : Exit Sub
        If TreeView1.SelectedNode.Parent Is Nothing Then MsgBox("This page already in root") : Exit Sub
        Dim node = TreeView1.SelectedNode
        Dim parent_node = TreeView1.SelectedNode.Parent
        Dim nodes As TreeNodeCollection

        If parent_node.Parent Is Nothing Then nodes = TreeView1.Nodes Else nodes = parent_node.Parent.Nodes
        Dim ind = nodes.IndexOf(parent_node)
        node.Remove()
        nodes.Insert(ind, node)
        TreeView1.SelectedNode = node
    End Sub
    'getNewName - CheckAvailName
    Private Function getNewName(node As TreeNode, name As String, Optional checkInsideNode As Boolean = False)
        Dim c As Integer = 0
        Dim new_name As String = name

        If Not checkInsideNode Then
            If node Is Nothing OrElse node.Parent Is Nothing Then
                Do While TreeView1.Nodes(new_name) IsNot Nothing
                    c += 1
                    new_name = name + " (" + c.ToString + ")"
                Loop
            ElseIf node.Parent IsNot Nothing Then
                Do While node.Parent.Nodes(new_name) IsNot Nothing
                    c += 1
                    new_name = name + " (" + c.ToString + ")"
                Loop
            End If
        Else
            Do While node.Nodes(new_name) IsNot Nothing
                c += 1
                new_name = name + " (" + c.ToString + ")"
            Loop
        End If

        Return new_name
    End Function

    'Save
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim t = DateTime.Now
        ToolStripButton22.Visible = True : Me.Refresh()
        order.Clear()
        If Not DirectoryExists(SAVE_PATH) Then CreateDirectory(SAVE_PATH)
        Dim forceResaveAll As Boolean = False
        If SAVE_PATH <> SAVE_PATH_START Then forceResaveAll = True
        For Each node As TreeNode In TreeView1.Nodes
            saveNodeRecursively(node, forceResaveAll)
        Next

        'Save order
        Dim w As New StreamWriter(SAVE_PATH + "\!order.ord")
        For Each o In order
            w.WriteLine(o)
        Next
        w.Close()

        ToolStripButton22.Visible = False

        Dim t2 = DateTime.Now - t
        'Label7.Text = "Last save time: " + t2.Seconds.ToString + "." + t2.Milliseconds.ToString + " seconds"
        Label7.Text = "Last save time: " + Math.Round(t2.TotalSeconds, 3).ToString + " seconds"
    End Sub
    Private Sub saveNodeRecursively(node As TreeNode, Optional forceResaveAll As Boolean = False)
        Dim dir = ""
        Dim p = node.FullPath
        order.Add(p)
        If p.Contains("\") Then dir = p.Substring(0, p.LastIndexOf("\"))
        If Not DirectoryExists(SAVE_PATH + "\" + dir) Then CreateDirectory(SAVE_PATH + "\" + dir)

        If forceResaveAll And content(node).Contains("{%%%UNLOADED%%%") Then
            Dim f = SAVE_PATH_START + "\" + node.FullPath + ".zam"
            If FileExists(f) Then content(node) = load_content(f, node)
        End If

        If Not content(node).Contains("{%%%UNLOADED%%%") Then
            Dim f = SAVE_PATH + "\" + dir + "\" + node.Text + ".zam"
            Dim w As New StreamWriter(f, False, System.Text.Encoding.UTF8)
            'w.AutoFlush = True
            If content_needCrypt.Contains(node) Then
                w.Write("CRYPF:" + crypt.EncryptData(content(node)))
            ElseIf EncryptContentToolStripMenuItem.Checked Then
                w.Write("CRYPR:" + crypt.EncryptData(content(node)))
            Else
                w.Write(content(node))
            End If
            w.Flush()
            w.Close()
        End If


        If node.ForeColor.Name <> "0" Or node.BackColor.Name <> "0" Or content_bgcolor.ContainsKey(node) Then
            Dim colStr = ""

            Dim c As Color = node.ForeColor
            colStr = c.A.ToString + ";" + c.R.ToString + ";" + c.G.ToString + ";" + c.B.ToString + "\"

            c = node.BackColor
            colStr += c.A.ToString + ";" + c.R.ToString + ";" + c.G.ToString + ";" + c.B.ToString + "\"

            If content_bgcolor.ContainsKey(node) Then
                c = content_bgcolor(node)
            Else
                c = Color.White
            End If
            colStr += c.R.ToString + ";" + c.G.ToString + ";" + c.B.ToString

            Dim w1 As New StreamWriter(SAVE_PATH + "\" + dir + "\" + node.Text + ".bcg")
            w1.WriteLine(colStr)
            w1.Close()
        Else
            If FileExists(SAVE_PATH + "\" + dir + "\" + node.Text + ".bcg") Then
                DeleteFile(SAVE_PATH + "\" + dir + "\" + node.Text + ".bcg")
            End If
        End If

        For Each node_sub As TreeNode In node.Nodes
            saveNodeRecursively(node_sub, forceResaveAll)
        Next
    End Sub

#Region "Toolbar"
    'Text formating - set page color
    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            Dim cd As New ColorDialog
            cd.ShowDialog()

            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            rtf.BackColor = cd.Color
            Dim node = content_associatedNodes(rtf)
            If cd.Color = Color.White Then
                If content_bgcolor.ContainsKey(node) Then content_bgcolor.Remove(node)
            Else
                If content_bgcolor.ContainsKey(node) Then
                    content_bgcolor(node) = cd.Color
                Else
                    content_bgcolor.Add(node, cd.Color)
                End If
            End If
        End If
    End Sub
    'Text formating - set font
    Private Sub ToolStripComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox2.TextChanged
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            Dim f = rtf.SelectionFont
            If f IsNot Nothing Then
                Try
                    rtf.SelectionFont = New Font(New FontFamily(ToolStripComboBox2.Text), CSng(ToolStripComboBox1.Text), f.Style)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                rtf.SelectionFont = New Font(New FontFamily(ToolStripComboBox2.Text), CSng(ToolStripComboBox1.Text), New FontStyle)
            End If
            rtf.Select()
        End If
    End Sub
    'Text formating - font size
    Private Sub ToolStripComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.TextChanged
        If refr Then Exit Sub

        'Replace decimal separator, and keep selection
        Dim selOrig = ToolStripComboBox1.SelectionStart
        Dim sep = Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
        ToolStripComboBox1.Text = ToolStripComboBox1.Text.Replace(",", sep).Replace(".", sep)
        ToolStripComboBox1.SelectionStart = selOrig

        If Not IsNumeric(ToolStripComboBox1.Text) Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            Dim f = rtf.SelectionFont
            If f Is Nothing Then f = New Font(New FontFamily(ToolStripComboBox2.Text), CSng(ToolStripComboBox1.Text))
            rtf.SelectionFont = New Font(f.FontFamily, CSng(ToolStripComboBox1.Text), f.Style)
            'rtf.Select()
        End If
    End Sub
    'Text formating - font color
    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            Dim cd As New ColorDialog
            cd.ShowDialog()

            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            rtf.SelectionColor = cd.Color
        End If
    End Sub
    'Text formating - font back color
    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            Dim cd As New ColorDialog
            cd.ShowDialog()

            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            rtf.SelectionBackColor = cd.Color
        End If
    End Sub
    'Text formating - font bold
    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            ToolStripButton8.Checked = Not ToolStripButton8.Checked
            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            Dim f = rtf.SelectionFont
            Dim fs = f.Style
            If ToolStripButton8.Checked Then
                If Not f.Bold Then fs = fs + FontStyle.Bold
            Else
                If f.Bold Then fs = fs - FontStyle.Bold
            End If

            rtf.SelectionFont = New Font(f.FontFamily, CSng(ToolStripComboBox1.Text), fs)
            rtf.Select()
        End If
    End Sub
    'Text formating - font italic
    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            ToolStripButton9.Checked = Not ToolStripButton9.Checked
            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            Dim f = rtf.SelectionFont
            Dim fs = f.Style
            If ToolStripButton9.Checked Then
                If Not f.Italic Then fs = fs + FontStyle.Italic
            Else
                If f.Italic Then fs = fs - FontStyle.Italic
            End If

            rtf.SelectionFont = New Font(f.FontFamily, CSng(ToolStripComboBox1.Text), fs)
            rtf.Select()
        End If
    End Sub
    'Text formating - font underline
    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            ToolStripButton10.Checked = Not ToolStripButton10.Checked
            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            Dim f = rtf.SelectionFont
            Dim fs = f.Style
            If ToolStripButton10.Checked Then
                If Not f.Underline Then fs = fs + FontStyle.Underline
            Else
                If f.Underline Then fs = fs - FontStyle.Underline
            End If

            rtf.SelectionFont = New Font(f.FontFamily, CSng(ToolStripComboBox1.Text), fs)
            rtf.Select()
        End If
    End Sub
    'Text formating - font strikeout
    Private Sub ToolStripButton15_Click(sender As Object, e As EventArgs) Handles ToolStripButton15.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            ToolStripButton15.Checked = Not ToolStripButton15.Checked
            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            Dim f = rtf.SelectionFont
            Dim fs = f.Style
            If ToolStripButton15.Checked Then
                If Not f.Strikeout Then fs = fs + FontStyle.Strikeout
            Else
                If f.Strikeout Then fs = fs - FontStyle.Strikeout
            End If

            rtf.SelectionFont = New Font(f.FontFamily, CSng(ToolStripComboBox1.Text), fs)
            rtf.Select()
        End If
    End Sub
    'Text formating - align left
    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            If Not ToolStripButton11.Checked Then
                refr = True
                ToolStripButton11.Checked = True
                ToolStripButton12.Checked = False
                ToolStripButton13.Checked = False
                ToolStripButton14.Checked = False
                Dim rtf = tab.Controls.OfType(Of RichTextBox).First
                rtf.SelectionAlignment = HorizontalAlignment.Left
                refr = False
            End If
        End If
    End Sub
    'Text formating - align center
    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs) Handles ToolStripButton12.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            If Not ToolStripButton12.Checked Then
                refr = True
                ToolStripButton12.Checked = True
                ToolStripButton11.Checked = False
                ToolStripButton13.Checked = False
                ToolStripButton14.Checked = False
                Dim rtf = tab.Controls.OfType(Of RichTextBox).First
                rtf.SelectionAlignment = HorizontalAlignment.Center
                refr = False
            End If
        End If
    End Sub
    'Text formating - align right
    Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs) Handles ToolStripButton13.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            If Not ToolStripButton13.Checked Then
                refr = True
                ToolStripButton13.Checked = True
                ToolStripButton11.Checked = False
                ToolStripButton12.Checked = False
                ToolStripButton14.Checked = False
                Dim rtf = tab.Controls.OfType(Of RichTextBox).First
                rtf.SelectionAlignment = HorizontalAlignment.Right
                refr = False
            End If
        End If
    End Sub
    'Text formating - align justify
    Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs) Handles ToolStripButton14.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            If Not ToolStripButton14.Checked Then
                refr = True
                ToolStripButton14.Checked = True
                ToolStripButton11.Checked = False
                ToolStripButton12.Checked = False
                ToolStripButton13.Checked = False
                Dim rtf = tab.Controls.OfType(Of RichTextBox).First
                rtf.SelectionAlignment = HorizontalAlignment.Center
                refr = False
            End If
        End If
    End Sub
    'Text formating - insert image
    Private Sub ToolStripButton16_Click(sender As Object, e As EventArgs) Handles ToolStripButton16.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            Dim f As New OpenFileDialog
            f.Filter = "All images |*.jpg;*.jpeg;*.gif;*.png;*.bmp;"
            f.Multiselect = True
            Dim orgdata = Clipboard.GetDataObject
            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            If f.ShowDialog = DialogResult.OK Then
                For Each fname As String In f.FileNames
                    Dim img As Image = Image.FromFile(fname)
                    Clipboard.SetImage(img)
                    rtf.Paste()
                Next
            End If
        End If
    End Sub
    'Text formating - insert table
    Private Sub ToolStripButton17_Click(sender As Object, e As EventArgs) Handles ToolStripButton17.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab Is Nothing Then Exit Sub

        For y As Integer = 1 To 5
            For x As Integer = 1 To 5
                Panel1.Controls("Panel" + x.ToString + y.ToString).BackColor = Color.White
            Next
        Next
        Panel1.Parent = Me
        Panel1.Top = 40
        Panel1.Left = 550
        Panel1.BringToFront()
        Panel1.Visible = True
        Panel1.Select()
    End Sub
    'Force Encrypt
    Private Sub ToolStripButton25_Click(sender As Object, e As EventArgs) Handles ToolStripButton25.Click
        If refr Then Exit Sub
        ToolStripButton25.Checked = Not ToolStripButton25.Checked
        If ToolStripButton25.Checked Then
            ToolStripButton25.Image = ImageList1.Images(1)
        Else
            ToolStripButton25.Image = ImageList1.Images(0)
        End If

        Dim ind = TabControl1.SelectedIndex
        If ind < 0 Then
            ToolStripButton25.Checked = False
            ToolStripButton25.Enabled = False
            ToolStripButton25.Image = ImageList1.Images(0)
        Else
            If TabControl1.TabPages(ind).Text = "Search" Then
                ToolStripButton25.Checked = False
                ToolStripButton25.Enabled = False
                ToolStripButton25.Image = ImageList1.Images(0)
            ElseIf TabControl1.TabPages(ind).Controls.Count = 1 Then
                ToolStripButton25.Enabled = True
                Dim rtf = TabControl1.TabPages(ind).Controls.OfType(Of RichTextBox).First
                Dim node = content_associatedNodes(rtf)
                If ToolStripButton25.Checked AndAlso Not content_needCrypt.Contains(node) Then content_needCrypt.Add(node)
                If Not ToolStripButton25.Checked AndAlso content_needCrypt.Contains(node) Then content_needCrypt.Remove(node)
            End If
        End If
    End Sub
    'Move tab left
    Private Sub ToolStripButton30_Click(sender As Object, e As EventArgs) Handles ToolStripButton30.Click
        If TabControl1.TabPages.Count > 0 Then
            If TabControl1.SelectedTab IsNot Nothing Then
                If TabControl1.SelectedTab IsNot optionsTabPage Then
                    Dim ind As Integer = TabControl1.SelectedIndex
                    If ind > 0 Then
                        Dim t As TabPage = TabControl1.SelectedTab
                        TabControl1.TabPages.Remove(t)
                        TabControl1.TabPages.Insert(ind - 1, t)
                        TabControl1.SelectedTab = t
                    End If
                End If
            End If
        End If
    End Sub
    'Move tab right
    Private Sub ToolStripButton29_Click(sender As Object, e As EventArgs) Handles ToolStripButton29.Click
        If TabControl1.TabPages.Count > 0 Then
            If TabControl1.SelectedTab IsNot Nothing Then
                If TabControl1.SelectedTab IsNot optionsTabPage Then
                    Dim ind As Integer = TabControl1.SelectedIndex
                    If ind < TabControl1.TabPages.Count - 1 Then
                        Dim t As TabPage = TabControl1.SelectedTab
                        TabControl1.TabPages.Remove(t)
                        TabControl1.TabPages.Insert(ind + 1, t)
                        TabControl1.SelectedTab = t
                    End If
                End If
            End If
        End If
    End Sub

    'Text formating - refresh controls
    Private Sub rtf_SelectionChanged(sender As Object, e As EventArgs)
        Dim refrWasTrue As Boolean = refr
        refr = True
        Dim rtf = DirectCast(sender, RichTextBox)
        Dim f = rtf.SelectionFont
        If f IsNot Nothing Then
            ToolStripComboBox1.Text = f.Size.ToString
            ToolStripComboBox2.Text = f.Name
            If f.Bold Then ToolStripButton8.Checked = True Else ToolStripButton8.Checked = False
            If f.Italic Then ToolStripButton9.Checked = True Else ToolStripButton9.Checked = False
            If f.Underline Then ToolStripButton10.Checked = True Else ToolStripButton10.Checked = False
            If f.Strikeout Then ToolStripButton15.Checked = True Else ToolStripButton15.Checked = False
        End If
        If rtf.SelectionAlignment = HorizontalAlignment.Left Then
            ToolStripButton11.Checked = True
            ToolStripButton12.Checked = False
            ToolStripButton13.Checked = False
            ToolStripButton14.Checked = False
        End If
        If rtf.SelectionAlignment = HorizontalAlignment.Center Then
            ToolStripButton11.Checked = False
            ToolStripButton12.Checked = True
            ToolStripButton13.Checked = False
            ToolStripButton14.Checked = False
        End If
        If rtf.SelectionAlignment = HorizontalAlignment.Right Then
            ToolStripButton11.Checked = False
            ToolStripButton12.Checked = False
            ToolStripButton13.Checked = True
            ToolStripButton14.Checked = False
        End If
        If Not refrWasTrue Then refr = False
    End Sub
    'Options
    Private Sub ToolStripButton27_Click(sender As Object, e As EventArgs) Handles ToolStripButton27.Click
        If Not TabControl1.TabPages.Contains(optionsTabPage) Then
            TabControl1.TabPages.Add(optionsTabPage)
            TabControl1.SelectedTab = optionsTabPage
        Else
            TabControl1.SelectedTab = optionsTabPage
        End If
    End Sub
#End Region

#Region "Search"
    'search - in a note
    Private Sub ToolStripButton18_Click(sender As Object, e As EventArgs) Handles ToolStripButton18.Click
        If refr Then Exit Sub
        Dim tab = TabControl1.SelectedTab
        If tab IsNot Nothing Then
            Dim rtf = tab.Controls.OfType(Of RichTextBox).First
            Dim s = InputBox("What to search?")
            If s.Trim = "" Then Exit Sub
            rtf.Find(s)
        End If
    End Sub
    'search - global
    Dim tmpNodeList As New List(Of TreeNode)
    Private Sub ToolStripButton19_Click(sender As Object, e As EventArgs) Handles ToolStripButton19.Click
        Dim search = InputBox("What to search?")
        If search.Trim = "" Then Exit Sub

        TabControl1.TabPages.Add("Search")
        Dim tabpage = TabControl1.TabPages(TabControl1.TabPages.Count - 1)
        TabControl1.SelectedIndex = TabControl1.TabPages.Count - 1

        Dim rtf As New RichTextBox
        rtf.ReadOnly = True
        rtf.Dock = DockStyle.Fill
        rtf.HideSelection = False
        tabpage.Controls.Add(rtf)

        get_nodes_recursively()
        For Each node In tmpNodeList
            Dim added_lines As New List(Of Integer)
            Dim added As Boolean = False
            Dim start As Integer = 0
            Dim searchCount As Integer = 0
            Dim rtfTmp As New RichTextBox
            rtfTmp.WordWrap = False
            If content(node).Contains("{%%%UNLOADED%%%") Then
                Dim f = SAVE_PATH + "\" + node.FullPath + ".zam"
                If FileExists(f) Then content(node) = load_content(f, node)
            End If
            If content(node).ToUpper.StartsWith("{\RTF") Then
                rtfTmp.Rtf = content(node)
            Else
                rtfTmp.Text = content(node)
            End If

            Do While rtfTmp.Find(search, start, RichTextBoxFinds.None) <> -1
                If Not added Then
                    rtf.Text += node.FullPath + vbCrLf
                    rtf.Text += New String("-", node.Text.Length) + vbCrLf
                    added = True
                End If

                start = rtfTmp.Find(search, start, RichTextBoxFinds.None)

                'get row
                Dim row = rtfTmp.GetLineFromCharIndex(start)
                Dim text = rtfTmp.Lines(row)
                start += 1
                If added_lines.Contains(row) Then Continue Do Else added_lines.Add(row)

                'get previous row
                If row >= 1 Then text = rtfTmp.Lines(row - 1) + vbCrLf + text

                'get next row
                If row < rtfTmp.Lines.Count - 1 Then text = text + vbCrLf + rtfTmp.Lines(row + 1)

                rtf.Text += text + vbCrLf + vbCrLf
                searchCount += 1
                If searchCount > 100 Then Exit Do
            Loop
            If searchCount > 100 Then rtf.Text = rtf.Text + vbCrLf + "Maximum search result (100) reached" : Exit For
            If added Then rtf.Text += vbCrLf
        Next

        Dim startIndex = 0, index = 0
        Do While rtf.Text.IndexOf(search, startIndex) >= 0
            index = rtf.Text.IndexOf(search, startIndex)
            rtf.Select(index, search.Length)
            rtf.SelectionColor = Color.Yellow
            startIndex = index + search.Length
        Loop
        If rtf.Text.Length > 3 Then rtf.Select(rtf.Text.Length - 1, 0)
    End Sub
    Private Sub get_nodes_recursively(Optional node As TreeNode = Nothing)
        If node Is Nothing Then
            tmpNodeList.Clear()
            For Each sub_node In TreeView1.Nodes
                tmpNodeList.Add(sub_node)
                get_nodes_recursively(sub_node)
            Next
        Else
            For Each sub_node In node.Nodes
                tmpNodeList.Add(sub_node)
                get_nodes_recursively(sub_node)
            Next
        End If
    End Sub

#End Region

#Region "Main menu"
    'Menu File/Export current
    Private Sub CurrentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CurrentToolStripMenuItem.Click
        Dim tab = TabControl1.SelectedTab
        If tab Is Nothing Then MsgBox("No page selected") : Exit Sub
        Dim rtf = tab.Controls.OfType(Of RichTextBox).First
        Dim f As New SaveFileDialog
        f.Filter = "Text files (*.txt)|*.txt"
        If f.ShowDialog <> DialogResult.OK Then Exit Sub
        Dim w As New StreamWriter(f.FileName)
        w.Write(rtf.Text.Replace(vbLf, vbCrLf))
        w.Close()
    End Sub
    'Menu File/Export all
    Private Sub AllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AllToolStripMenuItem.Click
        Dim f As New FolderBrowserDialog
        If f.ShowDialog() <> DialogResult.OK Then Exit Sub

        For Each node As TreeNode In TreeView1.Nodes
            exportRecur(node, f.SelectedPath)
        Next
    End Sub
    Private Sub exportRecur(node As TreeNode, rootPath As String)
        Dim dir = ""
        Dim p = node.FullPath
        If p.Contains("\") Then dir = p.Substring(0, p.LastIndexOf("\"))
        If Not DirectoryExists(rootPath + "\" + dir) Then CreateDirectory(rootPath + "\" + dir)
        Dim f = rootPath + "\" + dir + "\" + node.Text + ".txt"
        Dim w As New StreamWriter(f)

        Dim rtf As New RichTextBox
        If content(node).ToUpper.StartsWith("{\RTF") Then
            rtf.Rtf = content(node)
        Else
            rtf.Text = content(node)
        End If
        w.Write(rtf.Text.Replace(vbLf, vbCrLf))
        w.Close()

        For Each node_sub As TreeNode In node.Nodes
            exportRecur(node_sub, rootPath)
        Next
    End Sub

    'Edit/Set font size for opened pages
    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click,
            ToolStripMenuItem6.Click, ToolStripMenuItem7.Click, ToolStripMenuItem8.Click, ToolStripMenuItem9.Click,
            ToolStripMenuItem10.Click, ToolStripMenuItem11.Click, ToolStripMenuItem12.Click, ToolStripMenuItem13.Click,
            ToolStripMenuItem14.Click, ToolStripMenuItem15.Click, ToolStripMenuItem16.Click

        Dim size = CInt(DirectCast(sender, ToolStripMenuItem).Text)
        For Each tabpage As TabPage In TabControl1.TabPages
            If tabpage Is optionsTabPage Then Continue For
            Dim rtf = tabpage.Controls.OfType(Of RichTextBox).First
            rtf.SelectionStart = 0
            rtf.SelectionLength = rtf.TextLength

            Dim f = rtf.Font
            rtf.SelectionFont = New Font(f.FontFamily, size, f.Style)
        Next
    End Sub

    'Menu File/Exit
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub
    'Menu Options/Encrypt
    Private Sub EncryptContentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EncryptContentToolStripMenuItem.Click
        If refr Then Exit Sub
        EncryptContentToolStripMenuItem.Checked = Not EncryptContentToolStripMenuItem.Checked
        ini.IniWriteValue("Main", "Encrypt", EncryptContentToolStripMenuItem.Checked.ToString)
    End Sub
    'Menu Options/Autosave on exit
    Private Sub AutosaveOnExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutosaveOnExitToolStripMenuItem.Click
        If refr Then Exit Sub
        AutosaveOnExitToolStripMenuItem.Checked = Not AutosaveOnExitToolStripMenuItem.Checked
        ini.IniWriteValue("Main", "Autosave", AutosaveOnExitToolStripMenuItem.Checked.ToString)
    End Sub
    'Menu Options/Remember last pos
    Private Sub RememberLastPositionAndSizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RememberLastPositionAndSizeToolStripMenuItem.Click
        If refr Then Exit Sub
        RememberLastPositionAndSizeToolStripMenuItem.Checked = Not RememberLastPositionAndSizeToolStripMenuItem.Checked
        ini.IniWriteValue("Main", "Remember_last_pos_and_size", RememberLastPositionAndSizeToolStripMenuItem.Checked.ToString)
    End Sub
    'Menu Options/Tray/Use
    Private Sub UseTrayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UseTrayToolStripMenuItem.Click
        UseTrayToolStripMenuItem.Checked = Not UseTrayToolStripMenuItem.Checked
        If UseTrayToolStripMenuItem.Checked Then
            StartInTrayToolStripMenuItem.Enabled = True
            MinimizeToTrayToolStripMenuItem.Enabled = True
            CloseToTrayToolStripMenuItem.Enabled = True
            main_class.tray.visible = True
            ini.IniWriteValue("Main", "Tray_Use", "True")
        Else
            StartInTrayToolStripMenuItem.Enabled = False
            MinimizeToTrayToolStripMenuItem.Enabled = False
            CloseToTrayToolStripMenuItem.Enabled = False
            main_class.tray.visible = False
            ini.IniWriteValue("Main", "Tray_Use", "False")
        End If
        tray_use = UseTrayToolStripMenuItem.Checked
    End Sub
    'Menu Options/Tray/Start
    Private Sub StartInTrayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartInTrayToolStripMenuItem.Click
        StartInTrayToolStripMenuItem.Checked = Not StartInTrayToolStripMenuItem.Checked
        ini.IniWriteValue("Main", "Tray_Start", StartInTrayToolStripMenuItem.Checked.ToString)
        tray_start = StartInTrayToolStripMenuItem.Checked
    End Sub
    'Menu Options/Tray/Minimize
    Private Sub MinimizeToTrayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MinimizeToTrayToolStripMenuItem.Click
        MinimizeToTrayToolStripMenuItem.Checked = Not MinimizeToTrayToolStripMenuItem.Checked
        ini.IniWriteValue("Main", "Tray_Min", MinimizeToTrayToolStripMenuItem.Checked.ToString)
    End Sub
    'Menu Options/Tray/Close
    Private Sub CloseToTrayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToTrayToolStripMenuItem.Click
        CloseToTrayToolStripMenuItem.Checked = Not CloseToTrayToolStripMenuItem.Checked
        ini.IniWriteValue("Main", "Tray_close", CloseToTrayToolStripMenuItem.Checked.ToString)
    End Sub
    'Menu Options/Backup/Create backup
    Private Sub CreateNowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateNowToolStripMenuItem.Click
        If Not DirectoryExists(SAVE_PATH + "\!!!Backups") Then CreateDirectory(SAVE_PATH + "\!!!Backups")

        Dim d = DateTime.Now
        Dim b = d.Year.ToString + "-" + d.Month.ToString + "-" + d.Day.ToString + " " + d.Hour.ToString + "." + d.Minute.ToString + "." + d.Second.ToString + "." + d.Millisecond.ToString
        CreateDirectory(SAVE_PATH + "\!!!Backups\" + b)

        For Each Dr In GetDirectories(SAVE_PATH)
            Dim dName = Dr.Substring(Dr.LastIndexOf("\") + 1)
            If Not dName.ToUpper = "!!!BACKUPS" Then
                CopyDirectory(Dr, SAVE_PATH + "\!!!Backups\" + b + "\" + dName)
            End If
        Next
        For Each f In GetFiles(SAVE_PATH)
            FileCopy(f, SAVE_PATH + "\!!!Backups\" + b + "\" + Path.GetFileName(f))
        Next
        BackupsToolStripMenuItem.DropDownItems.Add(b)
    End Sub
#End Region

#Region "Context Menu"
    Dim colorForContextMenu As Color
    Dim nodeForContextMenu As TreeNode = Nothing
    Dim nodeForContextMenu2 As TreeNode = Nothing
    'Open in new tab
    Private Sub OpenInNewTabToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenInNewTabToolStripMenuItem.Click
        'If TreeView1.SelectedNode IsNot Nothing Then openNote(TreeView1.SelectedNode, True)
        'openNewTabPage(TreeView1.SelectedNode)
        If nodeForContextMenu IsNot Nothing Then openNote(nodeForContextMenu, True)
        nodeForContextMenu = Nothing
    End Sub
    'Open all children
    Private Sub OpenAllChildrenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenAllChildrenToolStripMenuItem.Click
        If nodeForContextMenu IsNot Nothing Then
            For Each node As TreeNode In nodeForContextMenu.Nodes
                If node.Nodes.Count = 0 Then openNote(node, True)
            Next
        End If
        nodeForContextMenu = Nothing
    End Sub
    'Add
    Private Sub AddToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddToolStripMenuItem.Click
        If nodeForContextMenu IsNot Nothing Then
            nodeForContextMenu2 = nodeForContextMenu
            ToolStripButton2_Click(ToolStripButton2, New EventArgs)
            nodeForContextMenu = Nothing
            nodeForContextMenu2 = Nothing
        End If
    End Sub
    'Add to child
    Private Sub AddChildToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddChildToolStripMenuItem.Click
        If nodeForContextMenu IsNot Nothing Then
            nodeForContextMenu2 = nodeForContextMenu
            ToolStripButton3_Click(ToolStripButton3, New EventArgs)
            nodeForContextMenu = Nothing
            nodeForContextMenu2 = Nothing
        End If
    End Sub
    'Clone
    Private Sub CloneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloneToolStripMenuItem.Click
        If nodeForContextMenu IsNot Nothing Then
            nodeForContextMenu2 = nodeForContextMenu
            ToolStripButton21_Click(ToolStripButton21, New EventArgs)
            nodeForContextMenu = Nothing
            nodeForContextMenu2 = Nothing
        End If
    End Sub
    'Rename
    Private Sub RenameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameToolStripMenuItem.Click
        If nodeForContextMenu IsNot Nothing Then
            nodeForContextMenu2 = nodeForContextMenu
            ToolStripButton20_Click(ToolStripButton20, New EventArgs)
            nodeForContextMenu = Nothing
            nodeForContextMenu2 = Nothing
        End If
    End Sub
    'Delete
    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If nodeForContextMenu IsNot Nothing Then
            nodeForContextMenu2 = nodeForContextMenu
            ToolStripButton4_Click(ToolStripButton4, New EventArgs)
            nodeForContextMenu = Nothing
            nodeForContextMenu2 = Nothing
        End If
    End Sub
    'Copy as file
    Private Sub CopyAsFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyAsFileToolStripMenuItem.Click
        If nodeForContextMenu IsNot Nothing Then
            Dim file = (SAVE_PATH + "\" + nodeForContextMenu.FullPath + ".zam").Replace("\\", "\")
            If FileExists(file) Then
                Dim DataObject As New DataObject
                DataObject.SetData(DataFormats.FileDrop, True, {file})
                Clipboard.SetDataObject(DataObject)
            End If
        End If
    End Sub
    'Sort children
    Private Sub SortChildrensToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SortChildrensToolStripMenuItem.Click
        If nodeForContextMenu IsNot Nothing AndAlso nodeForContextMenu.Nodes.Count > 1 Then
            Dim g As New List(Of TreeNode)
            'g = nodeForContextMenu.Nodes.Cast(Of TreeNode).ToList
            For Each t As TreeNode In nodeForContextMenu.Nodes
                g.Add(t)
            Next
            g.Sort(New NodeSorter)
            For Each t As TreeNode In g
                t.Remove()
            Next
            For Each t As TreeNode In g
                nodeForContextMenu.Nodes.Add(t)
            Next
        End If
    End Sub
    'Set fore color
    Private Sub setForeColor(sender As Object, e As EventArgs)
        If nodeForContextMenu Is Nothing Then Exit Sub
        Dim t = DirectCast(sender, ToolStripMenuItem)
        If t.Text.ToUpper.StartsWith("CUSTOM") Then
            Dim cp As New ColorDialog
            If cp.ShowDialog <> DialogResult.OK Then Exit Sub
            nodeForContextMenu.ForeColor = cp.Color
        Else
            nodeForContextMenu.ForeColor = t.ForeColor
        End If
        nodeForContextMenu = Nothing
    End Sub
    'Set fore color ...
    Private Sub SetNoteForeColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetNoteForeColorToolStripMenuItem.Click
        If SetNoteForeColorToolStripMenuItem.Text.EndsWith("...") Then
            Dim cp As New ColorDialog
            If cp.ShowDialog <> DialogResult.OK Then Exit Sub
            nodeForContextMenu.ForeColor = cp.Color
        End If
    End Sub
    'Set back color
    Private Sub setBackColor(sender As Object, e As EventArgs)
        If nodeForContextMenu Is Nothing Then Exit Sub
        Dim t = DirectCast(sender, ToolStripMenuItem)
        If t.Text.ToUpper.StartsWith("CUSTOM") Then
            Dim cp As New ColorDialog
            If cp.ShowDialog <> DialogResult.OK Then Exit Sub
            nodeForContextMenu.BackColor = cp.Color
        Else
            nodeForContextMenu.BackColor = t.BackColor
        End If
        nodeForContextMenu = Nothing
    End Sub
    'Set back color ...
    Private Sub SetNoteBackColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetNoteBackColorToolStripMenuItem.Click
        If SetNoteBackColorToolStripMenuItem.Text.EndsWith("...") Then
            Dim cp As New ColorDialog
            If cp.ShowDialog <> DialogResult.OK Then Exit Sub
            nodeForContextMenu.BackColor = cp.Color
        End If
    End Sub
    'Past as new node
    Private Sub PastAsNewNoteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PastAsNewNoteToolStripMenuItem.Click
        pastAsNewNote()
        nodeForContextMenu = Nothing
    End Sub

    'Context Menu Hide
    Private Sub ContextMenuStrip1_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles ContextMenuStrip1.Closing
        nodeForContextMenu.BackColor = colorForContextMenu
    End Sub
#End Region

#Region "Table"
    Private Sub Panel1_Leave(sender As Object, e As EventArgs) Handles Panel1.Leave
        Panel1.Visible = False
    End Sub
    Private Sub PanelSub_MouseHover(sender As Object, e As EventArgs) Handles Panel11.MouseHover, Panel12.MouseHover,
        Panel13.MouseHover, Panel14.MouseHover, Panel15.MouseHover, Panel21.MouseHover, Panel22.MouseHover,
        Panel23.MouseHover, Panel24.MouseHover, Panel25.MouseHover, Panel31.MouseHover, Panel32.MouseHover,
        Panel33.MouseHover, Panel34.MouseHover, Panel35.MouseHover, Panel41.MouseHover, Panel42.MouseHover,
        Panel43.MouseHover, Panel44.MouseHover, Panel45.MouseHover, Panel51.MouseHover, Panel52.MouseHover,
        Panel53.MouseHover, Panel54.MouseHover, Panel55.MouseHover
        For y As Integer = 1 To 5
            For x As Integer = 1 To 5
                Panel1.Controls("Panel" + x.ToString + y.ToString).BackColor = Color.White
            Next
        Next

        Dim name As String = DirectCast(sender, Panel).Name
        Dim toX As String = name.Substring(name.Length - 1, 1)
        Dim toY As String = name.Substring(name.Length - 2, 1)
        For y As Integer = 1 To CInt(toY)
            For x As Integer = 1 To CInt(toX)
                Panel1.Controls("Panel" + y.ToString + x.ToString).BackColor = Color.Blue
            Next
        Next
    End Sub
    Private Sub PanelSub_Click(sender As Object, e As EventArgs) Handles Panel11.Click, Panel12.Click,
        Panel13.Click, Panel14.Click, Panel15.Click, Panel21.Click, Panel22.Click,
        Panel23.Click, Panel24.Click, Panel25.Click, Panel31.Click, Panel32.Click,
        Panel33.Click, Panel34.Click, Panel35.Click, Panel41.Click, Panel42.Click,
        Panel43.Click, Panel44.Click, Panel45.Click, Panel51.Click, Panel52.Click,
        Panel53.Click, Panel54.Click, Panel55.Click

        Dim tab = TabControl1.SelectedTab
        If tab Is Nothing Then Exit Sub
        Dim rtf = tab.Controls.OfType(Of RichTextBox).First

        'Dim tableRtf As New System.Text.StringBuilder
        'tableRtf.Append("{\rtf1")
        'tableRtf.Append("\trowd")
        'tableRtf.Append("\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs")
        'tableRtf.Append("\cellx1000")
        'tableRtf.Append("\trrh3000")
        'tableRtf.Append("\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs")
        'tableRtf.Append("\cellx3000")

        'tableRtf.Append("\intbl \cell \row")

        'tableRtf.Append("\pard")
        'tableRtf.Append("}")
        'rtf.SelectedRtf = tableRtf.ToString()
        'Exit Sub

        Dim cellWidth As Integer = 2000
        Dim name As String = DirectCast(sender, Panel).Name
        Dim toX As String = name.Substring(name.Length - 1, 1)
        Dim toY As String = name.Substring(name.Length - 2, 1)

        Dim sbTaRtf As New System.Text.StringBuilder
        sbTaRtf.Append("{\rtf1")
        For row As Integer = 1 To CInt(toY)
            sbTaRtf.Append("\trowd")
            For col As Integer = 1 To CInt(toX)
                sbTaRtf.Append("\cellx" + (cellWidth * col).ToString)
                'sbTaRtf.Append("\cellx1000") 'set that cell width to 1000
                'sbTaRtf.Append("\cellx2000")
                'sbTaRtf.Append("\cellx3000")
            Next
            sbTaRtf.Append("\intbl \cell \row")
        Next
        sbTaRtf.Append("\pard")
        sbTaRtf.Append("}")
        rtf.SelectedRtf = sbTaRtf.ToString()

        'Exit Sub

        'Dim cols As Integer = 5
        'Dim rows As Integer = 5
        'Dim cols_width As Integer = 200
        'Dim A As String = ""

        'A = "{\rtf1\ansi\ansicpg1252\deff0"
        'A += "{\fonttbl{\f0\froman\fprq2\fcharset0 Times New Roman;}}"
        'A += "\viewkind4\uc1\trowd\trqc\trgaph108\trleft-8"
        'A += "\trbrdrt\brdrs\brdrw10"
        'A += "\trbrdrl\brdrs\brdrw10"
        'A += "\trbrdrb\brdrs\brdrw10"
        'A += "\trbrdrr\brdrs\brdrw10"
        'For i As Integer = 1 To cols
        '    A += "\clbrdrt\brdrw15\brdrs"
        '    A += "\clbrdrl\brdrw15\brdrs"
        '    A += "\clbrdrb\brdrw15\brdrs"
        '    A += "\clbrdrr\brdrw15\brdrs"
        '    A += "\cellx"
        '    A += CStr(cols_width) + "\clbrdrt"
        'Next
        'A += A & "\pard\intbl\lang3082\f0\fs24"
        'For i As Integer = 1 To rows
        '    A += "\intbl\clmrg"
        '    For j As Integer = 1 To cols
        '        A += "\cell"
        '    Next
        '    A += "\row"
        'Next
        'A += "}"
        'rtf.SelectedRtf = A
    End Sub
#End Region

#Region "Options"
    'Dbl click action
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        Dim r = DirectCast(sender, RadioButton)
        If r.Checked Then
            If RadioButton1.Checked Then ini.IniWriteValue("Main", "Dbl_click_on_tab", "OpenNew")
            If RadioButton2.Checked Then ini.IniWriteValue("Main", "Dbl_click_on_tab", "Rename")
        End If
    End Sub
    'Default font size
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ini.IniWriteValue("Main", "Default_font_size", ComboBox1.SelectedItem.ToString)
    End Sub
    'Ask for page name when adding
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ini.IniWriteValue("Main", "Ask_for_page_name", CheckBox1.Checked.ToString)
    End Sub
    'order
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged, RadioButton4.CheckedChanged, RadioButton5.CheckedChanged
        Dim r = DirectCast(sender, RadioButton)
        If r.Checked Then
            If RadioButton3.Checked Then ini.IniWriteValue("Main", "Note_Order", "Keep")
            If RadioButton4.Checked Then ini.IniWriteValue("Main", "Note_Order", "Alphabetical")
            If RadioButton5.Checked Then ini.IniWriteValue("Main", "Note_Order", "noOrder")
        End If
    End Sub
    'Use images
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then TreeView1.ImageList = ImageList1 Else TreeView1.ImageList = Nothing
        ini.IniWriteValue("Main", "Use_Images", CheckBox2.Checked.ToString)
    End Sub
    'Set custom icon
    Private Sub SetCustomIconToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetCustomIconToolStripMenuItem.Click
        If TreeView1.SelectedNode Is Nothing Then Exit Sub
        Dim fb As New OpenFileDialog
        fb.Filter = "png files (*.png)|*.png"
        If fb.ShowDialog <> DialogResult.OK Then Exit Sub
        FileCopy(fb.FileName, SAVE_PATH + "\" + TreeView1.SelectedNode.FullPath + ".png")

        Dim image = Bitmap.FromFile(SAVE_PATH + "\" + TreeView1.SelectedNode.FullPath + ".png")
        ImageList1.Images.Add(SAVE_PATH + "\" + TreeView1.SelectedNode.FullPath + ".png", image)
        TreeView1.SelectedNode.ImageKey = SAVE_PATH + "\" + TreeView1.SelectedNode.FullPath + ".png"
        TreeView1.SelectedNode.SelectedImageKey = SAVE_PATH + "\" + TreeView1.SelectedNode.FullPath + ".png"
    End Sub
    'Don't load all at startup
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        ini.IniWriteValue("Main", "Not_Load_All", CheckBox3.Checked.ToString)
    End Sub
    'Delete confirmation
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        ini.IniWriteValue("Main", "Delete_Confirm", CheckBox4.Checked.ToString)
    End Sub
    'Path change
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        SAVE_PATH = TextBox1.Text.Trim
        ini.IniWriteValue("Main", "Notes_Path", TextBox1.Text.Trim)
    End Sub
    'Browse in options
    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged, RadioButton7.CheckedChanged, RadioButton8.CheckedChanged
        Dim r = DirectCast(sender, RadioButton)
        If r.Checked Then
            If RadioButton6.Checked Then ini.IniWriteValue("Main", "Brose_in_options", "left") : CheckBox5.Enabled = True
            If RadioButton7.Checked Then ini.IniWriteValue("Main", "Brose_in_options", "right") : CheckBox5.Enabled = True
            If RadioButton8.Checked Then ini.IniWriteValue("Main", "Brose_in_options", "replace") : CheckBox5.Enabled = False
        End If
    End Sub
    'Browse in options - close options
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        ini.IniWriteValue("Main", "Brose_in_options_close", CheckBox5.Checked.ToString)
    End Sub
    'Browse in search
    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged, RadioButton10.CheckedChanged, RadioButton11.CheckedChanged
        Dim r = DirectCast(sender, RadioButton)
        If r.Checked Then
            If RadioButton9.Checked Then ini.IniWriteValue("Main", "Brose_in_search", "left") : CheckBox6.Enabled = True
            If RadioButton10.Checked Then ini.IniWriteValue("Main", "Brose_in_search", "right") : CheckBox6.Enabled = True
            If RadioButton11.Checked Then ini.IniWriteValue("Main", "Brose_in_search", "replace") : CheckBox6.Enabled = False
        End If
    End Sub
    'Browse in search - close search
    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        ini.IniWriteValue("Main", "Brose_in_search_close", CheckBox6.Checked.ToString)
    End Sub
    'Use color preset
    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        ini.IniWriteValue("Main", "Use_Color_Presets", CheckBox7.Checked.ToString)
        If CheckBox7.Checked Then
            If SetNoteForeColorToolStripMenuItem.DropDownItems.Count = 0 Then
                For Each cName In colorArray
                    Dim y As New ToolStripMenuItem
                    y.Text = cName.ToString
                    y.ForeColor = Color.FromName(cName.ToString)
                    SetNoteForeColorToolStripMenuItem.DropDownItems.Add(y)
                    AddHandler y.Click, AddressOf setForeColor

                    y = New ToolStripMenuItem
                    y.Text = cName.ToString
                    y.BackColor = Color.FromName(cName.ToString)
                    SetNoteBackColorToolStripMenuItem.DropDownItems.Add(y)
                    AddHandler y.Click, AddressOf setBackColor
                Next
                SetNoteForeColorToolStripMenuItem.Text = "Set Note ForeColor"
                SetNoteBackColorToolStripMenuItem.Text = "Set Note BackColor"
            End If
        Else
            If SetNoteForeColorToolStripMenuItem.DropDownItems.Count > 0 Then
                For i As Integer = SetNoteForeColorToolStripMenuItem.DropDownItems.Count - 1 To 0 Step -1
                    SetNoteForeColorToolStripMenuItem.DropDownItems.RemoveAt(i)
                Next
                For i As Integer = SetNoteBackColorToolStripMenuItem.DropDownItems.Count - 1 To 0 Step -1
                    SetNoteBackColorToolStripMenuItem.DropDownItems.RemoveAt(i)
                Next
                SetNoteForeColorToolStripMenuItem.Text = "Set Note ForeColor ..."
                SetNoteBackColorToolStripMenuItem.Text = "Set Note BackColor ..."
            End If
        End If
    End Sub
    'Hot tracking
    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        ini.IniWriteValue("Main", "HotTrack", CheckBox8.Checked.ToString)
    End Sub
    'Focus rtf on tree click
    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        focusRtfNextTime = Nothing
        ini.IniWriteValue("Main", "HotTrack", CheckBox9.Checked.ToString)
    End Sub
#End Region

#Region "Drag n drop"
    Dim alt As Boolean = False
    Dim draggedNode As TreeNode = Nothing
    Private Sub TreeView1_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles TreeView1.ItemDrag
        draggedNode = DirectCast(e.Item, TreeNode)

        Dim DataObject As New DataObject
        Dim file = Path.Combine(Application.StartupPath, SAVE_PATH + "\" + draggedNode.FullPath + ".zam")
        If FileExists(file) Then DataObject.SetData(DataFormats.FileDrop, True, {file})
        TreeView1.DoDragDrop(DataObject, DragDropEffects.All)
    End Sub

    Private Sub TreeView1_DragEnter(sender As Object, e As DragEventArgs) Handles TreeView1.DragEnter
        If (e.KeyState And 8) = 8 Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.Move
        End If
        If (e.KeyState And 32) = 32 Then alt = True Else alt = False
    End Sub
    Private Sub TreeView1_DragOver(sender As Object, e As DragEventArgs) Handles TreeView1.DragOver
        If (e.KeyState And 8) = 8 Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.Move
        End If
        If (e.KeyState And 32) = 32 Then alt = True Else alt = False
    End Sub

    'Draw label
    Private Sub TreeView1_GiveFeedback(sender As Object, e As GiveFeedbackEventArgs) Handles TreeView1.GiveFeedback
        'e.UseDefaultCursors = False
        'Cursor.Current = Cursors.Cross

        'If ((e.Effect And DragDropEffects.Copy) = DragDropEffects.Copy) Then
        If (e.Effect = DragDropEffects.Copy Or e.Effect = DragDropEffects.Move) Then
            If e.Effect = DragDropEffects.Move Then
                If alt Then Label6.Text = "(ch pos) " + draggedNode.Text Else Label6.Text = draggedNode.Text
            Else
                If alt Then Label6.Text = "(ch pos)(copy) " + draggedNode.Text Else Label6.Text = "(copy) " + draggedNode.Text
            End If

            Label6.Visible = True
            Dim pt = TreeView1.PointToClient(Cursor.Position)
            Label6.Location = New Point(pt.X + 15, pt.Y + 5)
        Else
            Label6.Visible = False
        End If
    End Sub

    'Handle cancel with esc
    Private Sub TreeView1_QueryContinueDrag(sender As Object, e As QueryContinueDragEventArgs) Handles TreeView1.QueryContinueDrag
        'ESC pressed
        If e.EscapePressed Then
            e.Action = DragAction.Cancel
            Exit Sub
        End If
        'Drop!
        If e.KeyState = 0 Then
            e.Action = DragAction.Drop
            Exit Sub
        End If
        e.Action = DragAction.Continue
    End Sub

    'DROP
    Private Sub TreeView1_DragDrop(sender As Object, e As DragEventArgs) Handles TreeView1.DragDrop
        Label6.Visible = False

        If draggedNode IsNot Nothing Then
            Dim pt As Point = TreeView1.PointToClient(New Point(e.X, e.Y))
            Dim DestinationNode As TreeNode = TreeView1.GetNodeAt(pt)

            Dim old_path = SAVE_PATH + "\" + draggedNode.FullPath
            Dim new_name As String

            If DestinationNode IsNot Nothing Then
                If Not DestinationNode Is draggedNode Then
                    Dim DragPath = draggedNode.FullPath.ToUpper + "\"
                    Dim DestPath = DestinationNode.FullPath.ToUpper + "\"
                    If DestPath.StartsWith(DragPath) Then MsgBox("Move a node to its child node won't work.") : draggedNode = Nothing : Exit Sub

                    If e.Effect = DragDropEffects.Move Then draggedNode.Remove()
                    new_name = getNewName(DestinationNode, draggedNode.Text, True)
                    If e.Effect = DragDropEffects.Copy Then
                        Dim old_content = content(draggedNode)
                        draggedNode = New TreeNode With {.Name = new_name, .Text = new_name}
                        content.Add(draggedNode, old_content)
                    End If

                    If alt Then
                        Dim ind = DestinationNode.Index
                        Dim p As TreeNodeCollection
                        If DestinationNode.Parent Is Nothing Then p = TreeView1.Nodes Else p = DestinationNode.Parent.Nodes
                        p.Insert(ind, draggedNode)
                    Else
                        DestinationNode.Nodes.Add(draggedNode)
                        DestinationNode.Expand()
                    End If
                Else
                    draggedNode = Nothing : Exit Sub
                End If
            Else
                If e.Effect = DragDropEffects.Move Then draggedNode.Remove()
                new_name = getNewName(Nothing, draggedNode.Text)
                If e.Effect = DragDropEffects.Copy Then
                    Dim old_content = content(draggedNode)
                    draggedNode = New TreeNode With {.Name = new_name, .Text = new_name}
                    content.Add(draggedNode, old_content)
                End If

                TreeView1.Nodes.Add(draggedNode)
            End If
            draggedNode.Name = new_name
            draggedNode.Text = new_name

            If e.Effect = DragDropEffects.Move Then
                If Not Path.GetFullPath(old_path).ToUpper = Path.GetFullPath(SAVE_PATH + "\" + draggedNode.FullPath).ToUpper Then
                    If DirectoryExists(Path.GetDirectoryName(old_path)) Then
                        For Each file In GetFiles(Path.GetDirectoryName(old_path), FileIO.SearchOption.SearchTopLevelOnly, {Path.GetFileName(old_path) + ".*"})
                            MoveFile(file, SAVE_PATH + "\" + draggedNode.FullPath + Path.GetExtension(file))
                        Next
                    End If
                    If DirectoryExists(old_path) Then MoveDirectory(old_path, SAVE_PATH + "\" + draggedNode.FullPath)
                End If
            End If

            draggedNode = Nothing
        End If
    End Sub

#End Region
End Class
