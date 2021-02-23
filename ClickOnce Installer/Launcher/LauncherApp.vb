Module LauncherApp

    Sub Main()

        ' Get current path
        Dim MyPath As String = Environment.CurrentDirectory

        Try

            ' Launch VBScript file
            Dim scriptProc = New Process()
            With scriptProc
                .StartInfo.FileName = "wscript"
                .StartInfo.WorkingDirectory = MyPath
                .StartInfo.Arguments = " Application.vbs"
                .Start()
                .WaitForExit()
                .Close()
            End With

            ' Use shell to execute vbscript.
            'Shell("wscript """ & MyPath & "\Application.vbs""")

        Catch ex As Exception
            MsgBox("Error Launching Application." & vbCrLf & "Expecting 'Application.vbs' in the same folder." & vbCrLf & ex.Message.ToString)
        End Try
    End Sub

End Module
