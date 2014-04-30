Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.DataSource = System.Enum.GetValues(GetType(TextMessageColor))
        ComboBox1.SelectedIndex = 8
        ComboBox2.DataSource = System.Enum.GetValues(GetType(enumChatChannels))
        ComboBox2.SelectedIndex = 1
        ComboBox3.DataSource = System.Enum.GetValues(GetType(enumTalkTypes))

        Tibia_hWnd = FindWindow("tibiaclient", vbNullString)
        GetWindowThreadProcessId(Tibia_hWnd, Tibia_PID)
        Tibia_Handle = OpenProcess(PROCESS_ALL_ACCESS, False, Tibia_PID)
        Tibia_Base = Process.GetProcessById(Tibia_PID).MainModule.BaseAddress
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim TalkType As enumTalkTypes = CType(ComboBox3.SelectedValue, enumTalkTypes)
        SendMessage(TextBox1.Text, TalkType)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Color As TextMessageColor = CType(ComboBox1.SelectedValue, TextMessageColor)
        RecvMsg(Replace(TextBox2.Text, vbNewLine, vbLf), Color)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim Channel As enumChatChannels = CType(ComboBox2.SelectedValue, enumChatChannels)
        SendMsgChannel(TextBox3.Text, Channel)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SendPM(TextBox4.Text, TextBox5.Text)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dance()
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim Color As TextMessageColor = CType(ComboBox1.SelectedValue, TextMessageColor)
        RecvMsgUsingNETDLL(Replace(TextBox2.Text, vbNewLine, vbLf), Color)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub
End Class
