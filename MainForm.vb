Imports System.ComponentModel

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        defaultSettings()
    End Sub
    Sub addCheckboxData() 'Add fields to the checkboxs on main screen
        CheckedListBox1.Items.AddRange(GlobalVar.tests)
        CheckedListBox2.Items.AddRange(GlobalVar.protocols)
        CheckedListBox3.Items.AddRange(GlobalVar.temps)
    End Sub
    Sub defaultSettings() 'Fill in fields for testing
        addCheckboxData()
        TextBox1.Text = "2412"
        TextBox2.Text = "2442"
        TextBox3.Text = "2472"
        TextBox4.Text = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\Desktop\"
        TextBox5.Text = "2400"
        TextBox6.Text = "2483.5"
        TextBox8.Text = "31.24"
        TextBox9.Text = "31.16"
        TextBox10.Text = "31.2"
        CheckedListBox1.SetItemChecked(0, True)
        'CheckedListBox2.SetItemChecked(0, True)
        CheckedListBox3.SetItemChecked(1, True)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'Start program button
        If CheckBoxes() Then
            If isConnectedGPIB() Then
                If getTestTime() Then
                    Dim bgwProgress As New BackgroundWorker
                    Dim frmProgress As ProgressBar
                    ProgressBar.Show(Me)
                    frmProgress = ProgressBar
                    ' here on the main thread FormP references the default instance, 
                    ' which it wouldn't from another thread.
                    AddHandler bgwProgress.DoWork, AddressOf BGW_DoWork
                    bgwProgress.RunWorkerAsync(New Object() {frmProgress})
                End If
            Else
                MsgBox("Please Connect GPIB to Analyzer")
            End If
        End If
    End Sub
    Function CheckBoxes() 'Check if all fields have data in
        If Not TextBox1.Text.Equals("") Then
            If Not TextBox2.Text.Equals("") Then
                If Not TextBox3.Text.Equals("") Then
                    If Not TextBox5.Text.Equals("") Then
                        If Not TextBox6.Text.Equals("") Then
                                If Not TextBox8.Text.Equals("") Then
                                    If Not TextBox9.Text.Equals("") Then
                                        If Not TextBox10.Text.Equals("") Then
                                            If Not CheckedListBox1.CheckedItems.Count = 0 Then
                                                If Not CheckedListBox2.CheckedItems.Count = 0 Then
                                                    If Not CheckedListBox3.CheckedItems.Count = 0 Then
                                                        If Not TextBox4.Text.Equals("") Then
                                                            GlobalVar.freq(0) = TextBox1.Text
                                                            GlobalVar.freq(1) = TextBox2.Text
                                                            GlobalVar.freq(2) = TextBox3.Text
                                                            GlobalVar.location = TextBox4.Text
                                                            GlobalVar.psd(0) = TextBox5.Text + "MHz"
                                                            GlobalVar.psd(1) = TextBox6.Text + "MHz"
                                                            GlobalVar.obe(0) = (Convert.ToDecimal(TextBox5.Text) - 0.5).ToString
                                                            GlobalVar.obe(1) = (Convert.ToDecimal(TextBox6.Text) + 0.5).ToString
                                                            Return True
                                                        Else
                                                            MsgBox("Please Enter SaveLocation")
                                                            Return False
                                                        End If
                                                    Else
                                                        MsgBox("Please Select Temperatures")
                                                        Return False
                                                    End If
                                                Else
                                                    MsgBox("Please Select Protocols")
                                                    Return False
                                                End If
                                            Else
                                                MsgBox("Please Select Tests")
                                                Return False
                                            End If
                                        Else
                                            MsgBox("Please Enter Top Channel Offset")
                                            Return False
                                        End If
                                    Else
                                        MsgBox("Please Enter Middle Channel Offset")
                                        Return False
                                    End If
                                Else
                                    MsgBox("Please Enter Bottem Channel Offset")
                                    Return False
                                End If
                        Else
                            MsgBox("Please Enter Top Band Frequency")
                            Return False
                        End If
                    Else
                        MsgBox("Please Enter Bottem Band Frequency")
                        Return False
                    End If
                Else
                    MsgBox("Please Enter Top Channel Frequency")
                    Return False
                End If
            Else
                MsgBox("Please Enter Middle Channel Frequency")
                Return False
            End If
        Else
            MsgBox("Please Enter Bottem Channel Frequency")
            Return False
        End If
    End Function
    

    Private Sub Browse_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dialog As New FolderBrowserDialog
        If DialogResult.OK = dialog.ShowDialog Then
            TextBox4.Text = dialog.SelectedPath + "\"
        End If
    End Sub
    Function getTestTime()
        Dim psd = 0
        Dim obw = 0
        Dim obe = 0
        If CheckedListBox1.GetItemChecked(0) Then
            psd = CheckedListBox2.CheckedItems.Count * 3 * 135
        End If
        If CheckedListBox1.GetItemChecked(1) Then
            obw = CheckedListBox2.CheckedItems.Count * 2 * 60
        End If
        If CheckedListBox1.GetItemChecked(2) Then
            obe = (CheckedListBox2.CheckedItems.Count * 160 * 120) - 5540
        End If


        Dim iSpan As TimeSpan = TimeSpan.FromSeconds(Convert.ToDecimal((obe + obw + psd) / 10))
        Dim hours = iSpan.Hours.ToString.PadLeft(2, "0"c).ToString
        Dim mins = iSpan.Minutes.ToString.PadLeft(2, "0"c).ToString
        Dim secs = iSpan.Seconds.ToString.PadLeft(2, "0"c).ToString
        Dim time As String = hours.ToString + "h" + mins.ToString + "m" + secs.ToString
        If MessageBox.Show("This test will take " + time + "s" + vbNewLine + "Do you wish to continue ?", "TestTime", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub ComTest_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Sharp.isSharpConnected() Then
            Sharp.startWifi()
            Sharp.setChannel("bottem", "802.11b")
            Threading.Thread.Sleep(3000)
            Sharp.setChannel("middle", "802.11b")
            Threading.Thread.Sleep(3000)
            Sharp.setChannel("top", "802.11b")
            Threading.Thread.Sleep(3000)
            Sharp.setChannel("bottem", "802.11g")
            Threading.Thread.Sleep(3000)
            Sharp.setChannel("middle", "802.11g")
            Threading.Thread.Sleep(3000)
            Sharp.setChannel("top", "802.11g")
            Threading.Thread.Sleep(3000)
            Sharp.setChannel("bottem", "802.11n")
            Threading.Thread.Sleep(3000)
            Sharp.setChannel("middle", "802.11n")
            Threading.Thread.Sleep(3000)
            Sharp.setChannel("top", "802.11n")
            Threading.Thread.Sleep(3000)
            Sharp.stopWifi()
        Else
            MsgBox("No device connected")
        End If
        
    End Sub
    Private Sub BGW_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        Dim tempBool() As Boolean = {False, False, False}
        Dim tempTime = 0
        Dim frmProgress As ProgressBar
        'get the right instance of ProgressForm
        frmProgress = TryCast(e.Argument(0), ProgressBar)

        If isSharpConnected() Then
            GlobalVar.sharpConnected = True
        Else
            GlobalVar.sharpConnected = False
        End If

        OpenGPIB()

        If Me.InvokeRequired Then
            Me.Invoke(Sub()
                          Me.Visible = False
                          tempBool(0) = Me.CheckedListBox1.GetItemChecked(0)
                          tempBool(1) = Me.CheckedListBox1.GetItemChecked(1)
                          tempBool(2) = Me.CheckedListBox1.GetItemChecked(2)
                      End Sub)
        End If

        If GlobalVar.sharpConnected Then
            Sharp.startWifi()
        End If

        If tempBool(0) Then
            runPowerSpectralDensity(Me, frmProgress)
        End If
        If tempBool(1) Then
            runOccupiedBandwidth(Me, frmProgress)
        End If
        If tempBool(2) Then
            runOutBandEmmisions(Me, frmProgress)
        End If


        If GlobalVar.sharpConnected Then
            Sharp.stopWifi()
        End If

        MsgBox("Test Finished")
        frmProgress.CloseProgress_instanceSafe()

        If Me.InvokeRequired Then
            Me.Invoke(Sub()
                          Me.Visible = True
                      End Sub)
        End If

        Application.Restart()
    End Sub
End Class
