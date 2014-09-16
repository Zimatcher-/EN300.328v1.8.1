Module Serial
    Dim com As New IO.Ports.SerialPort
    Dim i As Integer = 0
    Sub openCom(comPort As String)
        com = My.Computer.Ports.OpenSerialPort(comPort)
        com.DataBits = 8
        com.Parity = IO.Ports.Parity.Even
        com.BaudRate = 115200
    End Sub
    Sub writeCom(text As String)
        com.WriteLine(text)
    End Sub
    Sub closeCom(comPort As String)
        com.Close()
    End Sub
End Module