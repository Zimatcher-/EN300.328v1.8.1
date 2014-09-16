Imports Ivi.Visa.Interop
Imports System.Xml
Imports System.IO

Module GPIB
    Public node As XmlNode
    Public device As String = ""
    Public GPIB_DeviceName As String()
    Dim doc As XmlDocument = New XmlDocument()
    Dim ioMgr As New ResourceManager
    Dim instrument As New FormattedIO488
    Dim idn As String

    Sub loadCommandsDoc() 'Loads the command document so it can be read from
        doc.Load(GlobalVar.commandsLocXML)
        node = doc.DocumentElement()
    End Sub
    Sub OpenGPIB()
        loadCommandsDoc()
        Dim temp As String()
        'Find all GPIB adresses
        temp = ioMgr.FindRsrc("GPIB?*INSTR")
        For k = 0 To temp.GetLength(0) - 1 Step 1
            System.Console.WriteLine(temp(k))
        Next
        'Substring string so you get GPIB id
        temp(0) = temp(0).Substring(7, 2)
        System.Console.WriteLine(temp(0))
        'Open connection with device
        instrument.IO = ioMgr.Open("GPIB0::" + temp(0))
        GPIB.instrument.IO.Timeout = 10000
        'Get device name
        instrument.WriteString("*IDN?")
        idn = instrument.ReadString()
        GPIB.instrument.IO.Timeout = 2000
        GPIB_DeviceName = idn.Split(",")
        device = GPIB_DeviceName(1)
        System.Console.WriteLine("Connected Spectrum Analyzer = " + GPIB_DeviceName(1))
        'Check if device is known device, if it does and commands to array
        For i = 0 To node.ChildNodes.Count - 1 Step 1
            If node.ChildNodes.ItemOf(i).ChildNodes.ItemOf(2).InnerText.Equals(GPIB_DeviceName(1)) Then
                ReDim GlobalVar.commands(0 To node.ChildNodes.ItemOf(0).ChildNodes.Count - 1)
                For j = 0 To node.ChildNodes.ItemOf(i).ChildNodes.Count - 1 Step 1
                    GlobalVar.commands(j) = node.ChildNodes.ItemOf(i).ChildNodes.ItemOf(j).InnerText
                    'System.Console.WriteLine(GlobalVar.commands(j))
                Next
            End If
        Next
    End Sub
    Sub setCenterFreq(freq As String)
        instrument.WriteString(GlobalVar.commands(3) + " " + freq)
    End Sub
    Sub setStartFreq(freq As String)
        instrument.WriteString(GlobalVar.commands(4) + " " + freq)
    End Sub
    Sub setStopFreq(freq As String)
        If device.Equals("FSQ-26") Then
            instrument.WriteString(GlobalVar.commands(5) + " " + freq)
        End If
    End Sub
    Sub setFreqSpan(freq As String)
            instrument.WriteString(GlobalVar.commands(6) + " " + freq)
    End Sub
    
    Sub setRfAtt(att As String)
            instrument.WriteString(GlobalVar.commands(7) + " " + att) '10 DB
    End Sub
    Sub setOffset(offset As String)
        instrument.WriteString(GlobalVar.commands(8) + " " + offset) '10dB
    End Sub
    Sub setRefLvl(refLvl As String)
            instrument.WriteString(GlobalVar.commands(9) + " " + refLvl) '10dBm
    End Sub
    Sub setResBand(resBand As String)
        If GlobalVar.commands(10).Contains(";") Then
            Dim temp = GlobalVar.commands(10).Split(";")
            instrument.WriteString(temp(0))
            instrument.WriteString(temp(1))
        Else
            instrument.WriteString(GlobalVar.commands(10))
        End If
        
    End Sub
    Sub setVidBand(vidBand As String)
        If GlobalVar.commands(11).Contains(";") Then
            Dim temp = GlobalVar.commands(10).Split(";")
            instrument.WriteString(temp(0))
            instrument.WriteString(temp(1))
        Else
            instrument.WriteString(GlobalVar.commands(11))
        End If
    End Sub
    Sub setContinuousSweep()
            instrument.WriteString(GlobalVar.commands(12))
    End Sub
    Sub setSingleSweep()
            instrument.WriteString(GlobalVar.commands(13))
    End Sub
    Sub setSweepTime(time As String)
            instrument.WriteString(GlobalVar.commands(14) + " " + time)
    End Sub
    Sub setSweepCount(sweCount As String)
            instrument.WriteString(GlobalVar.commands(15) + " " + sweCount) ' 1 
    End Sub
    Sub setSweepPoints(swePoints As String)
            instrument.WriteString(GlobalVar.commands(16) + " " + swePoints) ' 1 
    End Sub
    Sub setTraceType(type As String)
            instrument.WriteString(GlobalVar.commands(17) + " " + type)
    End Sub
    Sub setDetectorType(detector As String)
            instrument.WriteString(GlobalVar.commands(18) + " " + detector)
    End Sub
    Sub setFreeRun()
            instrument.WriteString(GlobalVar.commands(19))
    End Sub
    Sub setFilter(filter)
            instrument.WriteString(GlobalVar.commands(20) + " " + filter)
    End Sub
    Sub setTimeDomPower()
        instrument.WriteString(GlobalVar.commands(21))
    End Sub
    Sub setObwReading()
        instrument.WriteString(GlobalVar.commands(22))
    End Sub
    Sub marker1MaxPeak()
        instrument.WriteString(GlobalVar.commands(23))
    End Sub
    Sub preset()
        instrument.WriteString(GlobalVar.commands(24))
    End Sub
    Sub startMessurment()
        instrument.WriteString(GlobalVar.commands(25))
    End Sub
    Sub getScreenshot(freq As String, protocol As String, test As String, location As String)
        'Setting file full path name
        If Not My.Computer.FileSystem.DirectoryExists(location) Then
            My.Computer.FileSystem.CreateDirectory(location)
        End If
        'MsgBox(location)
        Dim fileName As String = protocol + "-" + freq + "-" + test + ".WMF"
        Dim fullPath As String = location + fileName
        'MsgBox(fullPath)

        'Store trace on device memory and start GPIB transfer
        If device.Equals("FSQ-26") Then
            instrument.WriteString("HCOP:DEV:LANG WMF")
            instrument.WriteString("HCOP:DEST 'MMEM'")
            instrument.WriteString("MMEM:NAME 'D:\temp.WMF'")
            instrument.WriteString("HCOP:ITEM:ALL")
            instrument.WriteString("HCOP:IMM")
            instrument.WriteString("FORM UINT,8")
            instrument.WriteString("MMEM:DATA? 'D:\temp.WMF'")
        ElseIf device.Equals("FSU") Then
            'Implement FSU Screenshot
        ElseIf device.Equals("PXA") Then
            instrument.WriteString("MMEM:STOR:SCR 'D:\temp.WMF'")
            instrument.WriteString("MMEM:DATA? 'D:\temp.WMF'")
        End If

        'Store GPIB data into byte array
        Dim BlockData As Byte() = instrument.ReadIEEEBlock(IEEEBinaryType.BinaryType_UI1, False, True)
        'Try writing byte array to file if cant error
        Try
            File.WriteAllBytes(fullPath, BlockData)
            System.Console.WriteLine("Saved Screenshot in " + fullPath)
        Catch ex As Exception
            MsgBox("Failed to write file")
        End Try
    End Sub
    Sub getTraceData(freq As String, protocol As String, test As String, location As String)
        'Setting file full path name
        'Dim location As String
        'location = (GlobalVar.location + (jobNo.Remove(6, 2) + "xx\") + (jobNo + "\") + "Testing_Logbooks\" + "Radio\" + "Lab Testing\" + "EN 300 328 v1.8.1 (WLAN)\" + protocol + "\" + test + "\")
        If Not My.Computer.FileSystem.DirectoryExists(location) Then
            My.Computer.FileSystem.CreateDirectory(location)
        End If
        Dim fileName As String = protocol + "-" + freq + "-" + test + ".DAT"
        Dim fullPath As String = location + fileName

        'Store trace on device memory and start GPIB transfer
        If device.Equals("FSQ-26") Then
            instrument.WriteString("HCOP:DEST 'MMEM'")
            instrument.WriteString("MMEM:STOR:TRAC 1, 'D:\test.DAT'")
            instrument.WriteString("FORM UINT,8")
            instrument.WriteString("MMEM:DATA? 'D:\test.DAT'")
        ElseIf device.Equals("FSU") Then
            'Implement FSU Screenshot
        ElseIf device.Equals("PXA") Then
            instrument.WriteString("MMEM:STOR:SCR 'D:\temp.DAT'")
            instrument.WriteString("MMEM:DATA? 'D:\temp.DAT'")
        End If

        GPIB.instrument.IO.Timeout = 10000
        'Store GPIB data into byte array
        Dim BlockData As Byte() = GPIB.instrument.ReadIEEEBlock(IEEEBinaryType.BinaryType_UI1, False, True)
        GPIB.instrument.IO.Timeout = 2000

        'Try writing byte array to file if cant error
        Try
            File.WriteAllBytes(fullPath, BlockData)
            System.Console.WriteLine("Saved Trace in " + fullPath)
        Catch ex As Exception
            MsgBox("Failed to write file")
        End Try
    End Sub
End Module
