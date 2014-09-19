Imports System.Management

Module Serial
    Dim com As New IO.Ports.SerialPort
    Dim comPort = ""
    Dim i As Integer = 0
    Function isDeviceConnected(device As String)
        Dim obj
        Try
            Dim searcher As New ManagementObjectSearcher( _
            "root\cimv2", _
            "SELECT * FROM Win32_SerialPort")

            For Each queryObj As ManagementObject In searcher.Get()
                obj = queryObj("Name")
                System.Console.WriteLine(obj.ToString)
                If obj.ToString.Contains("NB Series Serial Port") Then
                    Dim temp = obj.ToString.ToCharArray
                    comPort = temp(temp.Length - 5) + temp(temp.Length - 4) + temp(temp.Length - 3) + temp(temp.Length - 2)
                    'System.Console.WriteLine(comPort.ToString)
                End If
            Next

        Catch err As ManagementException
            MsgBox("An error occurred while querying for WMI data: ")
        End Try
        If Not comPort.Equals("") Then
            Return True
        Else
            Return False
        End If
    End Function
    function openCom()
        com = My.Computer.Ports.OpenSerialPort(comPort)
        com.DataBits = 8
        com.Parity = IO.Ports.Parity.Even
        com.BaudRate = 115200
    End Function
    Sub writeCom(text As String)
        If text.Contains(";") Then
            Dim tempText = text.Split(";")
            For i = 0 To tempText.Length - 1 Step 1
                com.WriteLine(tempText(i))
            Next
        Else
            com.WriteLine(text)
        End If
    End Sub
    Sub closeCom()
        com.Close()
    End Sub
End Module