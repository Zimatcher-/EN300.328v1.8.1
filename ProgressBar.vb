Public Class ProgressBar
    Dim time = 0
    Dim maxTime = 30
    Dim remainTime = 0
    Private Sub ProgressBar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.ProgressBar1.Minimum = 0
    End Sub

    Private Delegate Sub delegate_SetTimer(ByVal value As Boolean, ByVal max As Decimal, ByVal title As String)
    'setting the timer enabled and also setting max time
    Public Sub SetTimer_instanceSafe(ByVal value As Boolean, ByVal max As Decimal, ByVal title As String)
        If Me.InvokeRequired Then
            Me.Invoke(New delegate_SetTimer(AddressOf Me.SetTimer_instanceSafe), value, max, title)
        Else
            If value Then
                Me.Text = title
                Me.Timer1.Enabled = value
                Me.maxTime = max
                Me.time = 0
                Me.ProgressBar1.Maximum = max
            Else
                Me.Timer1.Enabled = value
            End If


        End If
    End Sub

    Private Delegate Sub delegate_PauseTimer()
    'setting the timer enabled and also setting max time
    Public Sub PauseTimer_instanceSafe()
        If Me.InvokeRequired Then
            Me.Invoke(New delegate_PauseTimer(AddressOf Me.PauseTimer_instanceSafe))
        Else
            Timer1.Enabled = Not Timer1.Enabled
        End If
    End Sub

    Private Delegate Sub delegate_CloseProgress()
    'closes the progressbar safely and exits the thread.
    Public Sub CloseProgress_instanceSafe()
        If Me.InvokeRequired Then
            Me.Invoke(New delegate_CloseProgress(AddressOf Me.CloseProgress_instanceSafe))
        Else
            Me.Close()
        End If
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        time = time + 1
        If time < maxTime Then
            remainTime = maxTime - time
        Else
            remainTime = 0
        End If
        'convert time to seconds, mins and hours.
        Dim iSpan As TimeSpan = TimeSpan.FromSeconds(Convert.ToDecimal(remainTime / 10))
        Dim hours = iSpan.Hours.ToString.PadLeft(2, "0"c).ToString
        Dim mins = iSpan.Minutes.ToString.PadLeft(2, "0"c).ToString
        Dim secs = iSpan.Seconds.ToString.PadLeft(2, "0"c).ToString
        Me.Label3.Text = ("Time Remaining : " + hours + "h, " + mins + "m, " + secs + "s")
        iSpan = TimeSpan.FromSeconds(Convert.ToDecimal(time / 10))
        hours = iSpan.Hours.ToString.PadLeft(2, "0"c).ToString
        mins = iSpan.Minutes.ToString.PadLeft(2, "0"c).ToString
        secs = iSpan.Seconds.ToString.PadLeft(2, "0"c).ToString
        Me.Label2.Text = ("Time Elapsed : " + hours + "h, " + mins + "m, " + secs + "s")
        'change progressbar value and text every timer tick.
        If time >= Me.ProgressBar1.Maximum Then
            Me.ProgressBar1.Value = Me.ProgressBar1.Maximum
        Else
            Me.ProgressBar1.Value = time
        End If

        Me.Label1.Text = (FormatNumber(((time / maxTime) * 100), 0).ToString + "%")
    End Sub
End Class