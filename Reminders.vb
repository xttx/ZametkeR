Public Class Reminders
    Public reminders As New Dictionary(Of Date, remind_struct)
    Public reminders_astralis As New Dictionary(Of TreeNode, remind_struct)
    Dim WithEvents timer As New Timer With {.Interval = 1000, .Enabled = True}

    Public Structure remind_struct
        Dim status As Boolean
        Dim msg As String
        Dim lineIndex As Integer
        Dim charcount As Integer
        Dim dt As DateTime
    End Structure

    Public Sub parse(txt As String)
        reminders.Clear()
        Dim l As Integer = 0
        Dim c As Integer = 0
        For Each line In txt.Split({CChar(vbLf)})
            Dim time As String = ""
            Dim dt As Date = Nothing
            Dim msg = ""

            Dim ind = line.IndexOf(" ")
            If ind < 0 Then
                'no space in line
                time = line.Trim
                If Not Date.TryParse(time, dt) Then l += 1 : Continue For
                c = time.Length
                msg = "Achtung!"
            Else
                'try parse space
                time = line.Substring(0, ind).Trim

                If Not Date.TryParse(time, dt) Then l += 1 : Continue For
                c = time.Length

                'try parse second space
                ind = line.IndexOf(" ", ind + 1)
                If ind > 2 Then
                    'try entire line
                    time = line.Trim
                    Dim dt2 As Date = Nothing
                    If Date.TryParse(time, dt2) Then
                        dt = dt2 : c = time.Length
                    Else
                        time = line.Substring(0, ind).Trim
                        dt2 = Nothing
                        If Date.TryParse(time, dt2) Then dt = dt2 : c = time.Length
                    End If
                End If

                msg = line.Substring(line.IndexOf(" ")).Trim
                If msg = "" Then msg = "Achtung!"
            End If

            Dim strct As remind_struct
            strct.msg = msg
            strct.lineIndex = l
            strct.charcount = c
            If dt > DateTime.Now Then
                strct.status = True
            Else
                strct.status = False
            End If

            If reminders.ContainsKey(dt) Then
                reminders(dt) = strct
            Else
                reminders.Add(dt, strct)
            End If
            l += 1
        Next
    End Sub

    Public Sub parse_astralis(node As TreeNode, txt As String)
        Dim lines = txt.Split({CChar(vbLf)}).ToList
        If lines.Count = 0 Then
            If reminders_astralis.ContainsKey(node) Then reminders_astralis.Remove(node)
        Else
            Dim dt As Date = Nothing
            If Not Date.TryParse(lines(0), dt) Then
                If reminders_astralis.ContainsKey(node) Then reminders_astralis.Remove(node)
                Exit Sub
            End If

            Dim strct As remind_struct
            strct.lineIndex = 0
            strct.charcount = lines(0).Length
            strct.dt = dt
            If dt > DateTime.Now Then
                strct.status = True
            Else
                strct.status = False
            End If

            Dim msg = ""
            If lines.Count = 1 Then
                msg = "Achtung!"
            Else
                lines.RemoveAt(0)
                For Each line In lines
                    msg += line + vbCrLf
                Next
            End If
            strct.msg = msg

            If reminders_astralis.ContainsKey(node) Then
                reminders_astralis(node) = strct
            Else
                reminders_astralis.Add(node, strct)
            End If
        End If
    End Sub

    Public Sub timer_tick() Handles timer.Tick
        Dim d = DateTime.Now
        For i As Integer = 0 To reminders.Keys.Count - 1
            Dim key = reminders.Keys(i)

            If reminders(key).status = True Then
                If d >= key And d <= key.AddSeconds(5) Then
                    Dim st = reminders(key)
                    st.status = False
                    reminders(key) = st

                    timer.Enabled = False
                    MsgBox(reminders(key).msg)
                    timer.Enabled = True
                End If
            End If
        Next
        For i As Integer = 0 To reminders_astralis.Keys.Count - 1
            Dim key = reminders_astralis.Keys(i)
            If reminders_astralis(key).status = True Then
                If d >= reminders_astralis(key).dt And d <= reminders_astralis(key).dt.AddSeconds(5) Then
                    Dim st = reminders_astralis(key)
                    st.status = False
                    reminders_astralis(key) = st

                    timer.Enabled = False
                    MsgBox(reminders_astralis(key).msg)
                    timer.Enabled = True
                End If
            End If
        Next
    End Sub
End Class
