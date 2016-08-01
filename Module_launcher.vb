Module Module_launcher
    Public SAVE_PATH As String = ".\Zametki"
    Public ini As New IniFileApi()
    Public tray_use As Boolean = False
    Public tray_start As Boolean = False
    Public main_class = New Main()
    Public loading_start As Date

    'Version 1
    '
    Public Sub Main()
        Application.EnableVisualStyles()
        Application.Run(main_class)
    End Sub

    'Version 2
    '
    'Public Sub Main(ByVal cmdArgs() As String)
    'End Sub

    'Version 3
    '
    'Public Function Main() As Integer
    'End Function

    'Version 4
    '
    'Public Function Main(ByVal cmdArgs() As String) As Integer
    'End Function

End Module
