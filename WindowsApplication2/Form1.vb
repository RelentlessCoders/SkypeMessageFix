Imports SKYPE4COMLib
'
'            RelentlessCoders
'       Skype Message Method/Bypass
'  
'  Basically the same as MRMURK4G3 version
Public Class Form1
    Dim oSkype As New SKYPE4COMLib.Skype
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sendMsg(ComboBox1.SelectedItem, TextBox1.Text) ' Calls the send message function with the values you have set
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oSkype.Attach() ' Attach to Skype via Skype4COMLib
        For Each user In oSkype.Friends
            ComboBox1.Items.Add(user.Handle) ' Add everyones username handle to the combobox
        Next

    End Sub

    Public Async Sub sendMsg(contact As String, msg As String)
        Clipboard.SetText(msg) ' Set your message to the copy clipboard
        Process.Start("skype:" + contact + "?chat") ' Open a chat with the selected user
        Await Task.Delay(500) ' Wait 500ms
        SendKeys.Send("^V") ' Paste the message copied to the clipboard
        Await Task.Delay(500) ' Wait 500ms
        SendKeys.Send("{ENTER}") ' Sends ENTER to send the message
    End Sub


End Class
