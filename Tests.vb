Imports System.Threading
Imports System.ComponentModel
Module Tests
    Dim ProgressBar As ProgressBar
    Sub runPowerSpectralDensity(form As Form1, progressBarForm As ProgressBar)
        preset()
        Dim maxPer As Integer = form.CheckedListBox2.CheckedItems.Count * 3

        Dim tempLocation As String
        If GlobalVar.sharpConnected Then
            MsgBox("Please make sure device is at Ambient")
        End If
        For i = 0 To form.CheckedListBox2.Items.Count - 1 Step 1
            If form.CheckedListBox2.GetItemChecked(i) = True Then
                tempLocation = (GlobalVar.location + "EN 300 328 v1.8.1 (WLAN)\" + form.CheckedListBox2.Items(i).ToString + "\Power Spectral Density\")
                If GlobalVar.sharpConnected Then
                    Sharp.setChannel("bottem", form.CheckedListBox2.Items(i).ToString)
                Else
                    MsgBox("Set device to channel frequency " + GlobalVar.freq(0) + "MHz on protocol " + form.CheckedListBox2.Items(i).ToString)
                End If
                GlobalVar.offset = form.TextBox8.Text + "dB"
                Threading.Thread.Sleep(500)
                'Threading.Thread.Sleep(10000)
                progressBarForm.SetTimer_instanceSafe(True, ((maxPer) * 135), "Power Spectral Density")
                psdReading(GlobalVar.freq(0), form.CheckedListBox2.CheckedItems.Item(i).ToString, tempLocation)
                progressBarForm.SetTimer_instanceSafe(False, ((maxPer) * 135), "Power Spectral Density")
                If GlobalVar.sharpConnected Then
                    Sharp.setChannel("middle", form.CheckedListBox2.Items(i).ToString)
                Else
                    MsgBox("Set device to channel frequency " + GlobalVar.freq(1) + "MHz on protocol " + form.CheckedListBox2.Items(i).ToString)
                End If
                GlobalVar.offset = form.TextBox9.Text + "dB"
                'Threading.Thread.Sleep(10000)
                progressBarForm.SetTimer_instanceSafe(True, ((maxPer) * 135), "Power Spectral Density")
                psdReading(GlobalVar.freq(1), form.CheckedListBox2.CheckedItems.Item(i).ToString, tempLocation)
                progressBarForm.SetTimer_instanceSafe(False, ((maxPer) * 135), "Power Spectral Density")
                If GlobalVar.sharpConnected Then
                    Sharp.setChannel("top", form.CheckedListBox2.Items(i).ToString)
                Else
                    MsgBox("Set device to channel frequency " + GlobalVar.freq(2) + "MHz on protocol " + form.CheckedListBox2.Items(i).ToString)
                End If
                GlobalVar.offset = form.TextBox10.Text + "dB"
                'Threading.Thread.Sleep(10000)
                progressBarForm.SetTimer_instanceSafe(True, ((maxPer) * 135), "Power Spectral Density")
                psdReading(GlobalVar.freq(2), form.CheckedListBox2.CheckedItems.Item(i).ToString, tempLocation)
                progressBarForm.SetTimer_instanceSafe(False, ((maxPer) * 135), "Power Spectral Density")
            End If
        Next
        progressBarForm.SetTimer_instanceSafe(False, 0, "Power Spectral Density")
    End Sub

    Sub psdReading(freq As String, protocol As String, tempLocation As String)
        setStartFreq(GlobalVar.psd(0))
        setStopFreq(GlobalVar.psd(1))
        setOffset(GlobalVar.offset)
        setRefLvl(GlobalVar.psd(2))
        setRfAtt(GlobalVar.psd(3))
        setResBand(GlobalVar.psd(4))
        setVidBand(GlobalVar.psd(5))
        setSweepTime(GlobalVar.psd(6))
        setSweepPoints(GlobalVar.psd(7))
        setDetectorType(GlobalVar.psd(8))
        setTraceType(GlobalVar.psd(9))
        setSingleSweep()
        setFreeRun()
        startMessurment()
        Threading.Thread.Sleep(10250)
        getScreenshot(freq + " MHz", protocol, "Power Spectral Density", tempLocation)
        getTraceData(freq + " MHz", protocol, "Power Spectral Density", tempLocation)
    End Sub
    Sub runOutBandEmmisions(form As Form1, progressBarForm As ProgressBar)
        preset()
        Dim maxPer As Integer = form.CheckedListBox2.CheckedItems.Count * 160
        Dim tempLocation As String = ""
        For i = 0 To form.CheckedListBox2.Items.Count - 1 Step 1
            If form.CheckedListBox2.GetItemChecked(i) Then
                For y = 0 To form.CheckedListBox3.Items.Count - 1 Step 1
                    If form.CheckedListBox3.GetItemChecked(y) Then
                        If GlobalVar.sharpConnected Then
                            MsgBox("Please set Temperature to " + form.CheckedListBox3.Items(y).ToString)
                        End If
                        For j = 0 To 1 Step 1
                            If j = 0 Then
                                If GlobalVar.sharpConnected Then
                                    Sharp.setChannel("bottem", form.CheckedListBox2.Items(i).ToString)
                                Else
                                    MsgBox("Set device to channel frequency " + GlobalVar.freq(0) + "MHz on protocol " + form.CheckedListBox2.Items(i).ToString + " and " + vbNewLine + "set temperture to " + form.CheckedListBox3.Items(y).ToString)
                                End If
                                GlobalVar.offset = form.TextBox8.Text + "dB"
                                tempLocation = (GlobalVar.location + "EN 300 328 v1.8.1 (WLAN)\" + form.CheckedListBox2.Items(i).ToString + "\Out Band Emission\" + form.CheckedListBox3.Items(y).ToString + "\" + "Bottem Channel (" + form.TextBox1.Text + "MHz)\")
                            ElseIf j = 1 Then
                                If GlobalVar.sharpConnected Then
                                    Sharp.setChannel("top", form.CheckedListBox2.Items(i).ToString)
                                Else
                                    MsgBox("Set device to channel frequency " + GlobalVar.freq(2) + "MHz on protocol " + form.CheckedListBox2.Items(i).ToString + " and " + vbNewLine + "set temperture to " + form.CheckedListBox3.Items(y).ToString)
                                End If
                                GlobalVar.offset = form.TextBox10.Text + "dB"
                                tempLocation = (GlobalVar.location + "EN 300 328 v1.8.1 (WLAN)\" + form.CheckedListBox2.Items(i).ToString + "\Out Band Emission\" + form.CheckedListBox3.Items(y).ToString + "\" + "Top Channel (" + form.TextBox3.Text + "MHz)\")
                            End If
                            progressBarForm.SetTimer_instanceSafe(True, ((maxPer) * 120) + 10, "Out Band Emissions")
                            For k = 0 To 39 Step 1
                                'System.Console.WriteLine((Convert.ToDecimal(GlobalVar.obe(0)) - (k * 1)).ToString)
                                obeReading((Convert.ToDecimal(GlobalVar.obe(0)) - (k * 1)).ToString, form.CheckedListBox2.Items(i).ToString, tempLocation)
                                'Threading.Thread.Sleep(500)
                            Next
                            For x = 0 To 39 Step 1
                                'System.Console.WriteLine((Convert.ToDecimal(GlobalVar.obe(1)) + (x * 1)).ToString)
                                obeReading((Convert.ToDecimal(GlobalVar.obe(1)) + (x * 1)).ToString, form.CheckedListBox2.Items(i).ToString, tempLocation)
                                'Threading.Thread.Sleep(500)
                            Next
                            progressBarForm.SetTimer_instanceSafe(False, ((maxPer) * 120) + 10, "Out Band Emissions")
                        Next
                    End If
                Next
            End If
        Next
        progressBarForm.SetTimer_instanceSafe(False, 0, "Out Band Emissions")
    End Sub
    Sub obeReading(freq As String, protocol As String, tempLocation As String)
        setCenterFreq(freq + "MHz")
        setFreqSpan(GlobalVar.obe(2))
        setOffset(GlobalVar.offset)
        setRefLvl(GlobalVar.obe(3))
        setRfAtt(GlobalVar.obe(4))
        setResBand(GlobalVar.obe(5))
        setVidBand(GlobalVar.obe(6))
        setSweepTime(GlobalVar.obe(7))
        setSweepPoints(GlobalVar.obe(8))
        setSweepCount(GlobalVar.obe(9))
        setDetectorType(GlobalVar.obe(10))
        setTraceType(GlobalVar.obe(11))
        setFilter(GlobalVar.obe(12))
        setSingleSweep()
        setFreeRun()
        setTimeDomPower()
        startMessurment()
        Threading.Thread.Sleep(11500)
        marker1MaxPeak()
        getScreenshot(freq + " MHz", protocol, "Out Band Emissions", tempLocation)
        'getTraceData(freq + " MHz", protocol, "Out Band Emissions", tempLocation)
    End Sub
    Sub runOccupiedBandwidth(form As Form1, progressBarForm As ProgressBar)
        preset()
        Dim maxPer As Integer = form.CheckedListBox2.CheckedItems.Count * 2

        Dim progValue = 0
        Dim tempLocation As String = ""
        If GlobalVar.sharpConnected Then
            MsgBox("Please make sure device is at Ambient")
        End If
        For i = 0 To form.CheckedListBox2.Items.Count - 1 Step 1
            If form.CheckedListBox2.GetItemChecked(i) Then
                tempLocation = (GlobalVar.location + "EN 300 328 v1.8.1 (WLAN)\" + form.CheckedListBox2.Items(i).ToString + "\Occupied Bandwidth\")
                For j = 0 To 1 Step 1
                    If j = 0 Then
                        If GlobalVar.sharpConnected Then
                            Sharp.setChannel("bottem", form.CheckedListBox2.Items(i).ToString)
                        Else
                            MsgBox("Set device to channel frequency " + GlobalVar.freq(0) + "MHz on protocol " + form.CheckedListBox2.Items(i).ToString)
                        End If
                        progressBarForm.SetTimer_instanceSafe(True, ((maxPer) * 60), "Occupied Bandwidth")
                        GlobalVar.offset = form.TextBox8.Text + "dB"
                        Threading.Thread.Sleep(500)
                        'Threading.Thread.Sleep(2000)
                        obwReading(GlobalVar.freq(0), form.CheckedListBox2.Items(i).ToString, tempLocation)
                        progressBarForm.SetTimer_instanceSafe(False, ((maxPer) * 60), "Occupied Bandwidth")
                    ElseIf j = 1 Then
                        If GlobalVar.sharpConnected Then
                            Sharp.setChannel("top", form.CheckedListBox2.Items(i).ToString)
                        Else
                            MsgBox("Set device to channel frequency " + GlobalVar.freq(2) + "MHz on protocol " + form.CheckedListBox2.Items(i).ToString)
                        End If
                        progressBarForm.SetTimer_instanceSafe(True, ((maxPer) * 60), "Occupied Bandwidth")
                        GlobalVar.offset = form.TextBox10.Text + "dB"
                        Threading.Thread.Sleep(500)
                        'Threading.Thread.Sleep(2000)
                        obwReading(GlobalVar.freq(2), form.CheckedListBox2.Items(i).ToString, tempLocation)
                        progressBarForm.SetTimer_instanceSafe(False, ((maxPer) * 60), "Occupied Bandwidth")
                    End If

                Next
            End If
        Next
        progressBarForm.SetTimer_instanceSafe(False, 0, "Occupied Bandwidth")
    End Sub
    Sub obwReading(freq As String, protocol As String, tempLocation As String)
        setCenterFreq(freq + "MHz")
        setFreqSpan(GlobalVar.obw(2))
        setOffset(GlobalVar.offset)
        setRefLvl(GlobalVar.obw(3))
        setRfAtt(GlobalVar.obw(4))
        setResBand(GlobalVar.obw(5))
        setVidBand(GlobalVar.obw(6))
        setSweepTime(GlobalVar.obw(7))
        setSweepPoints(GlobalVar.obw(8))
        setSweepCount(GlobalVar.obw(9))
        setDetectorType(GlobalVar.obw(10))
        setTraceType(GlobalVar.obw(11))
        setSingleSweep()
        setFreeRun()
        setObwReading()
        startMessurment()
        Threading.Thread.Sleep(4500)
        marker1MaxPeak()
        getScreenshot(freq + " MHz", protocol, "Occupied Bandwidth", tempLocation)
        'getTraceData(freq + " MHz", protocol, "Occupied Bandwidth", tempLocation)
    End Sub
End Module
