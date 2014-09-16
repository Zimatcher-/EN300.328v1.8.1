Module Sharp
    Sub startWifi(com As String)
        openCom(com)
        writeCom("WTFD 1")
    End Sub
    Sub setChannel(freq As String, protocol As String)
        If freq.Equals("bottem") Then
            If protocol.Equals("802.11b") Then
                writeCom("WIDL")
                writeCom("WBTX2 001021150000020")
            ElseIf protocol.Equals("802.11g") Then
                writeCom("WIDL")
                writeCom("WBTX2 001041200000020")
            ElseIf protocol.Equals("802.11n") Then
                writeCom("WIDL")
                writeCom("WBTX2 001121200000020")
            End If
        ElseIf freq.Equals("middle") Then
            If protocol.Equals("802.11b") Then
                writeCom("WIDL")
                writeCom("WBTX2 007021150000020")
            ElseIf protocol.Equals("802.11g") Then
                writeCom("WIDL")
                writeCom("WBTX2 007041200000020")
            ElseIf protocol.Equals("802.11n") Then
                writeCom("WIDL")
                writeCom("WBTX2 007121200000020")
            End If
        ElseIf freq.Equals("top") Then
            If protocol.Equals("802.11b") Then
                writeCom("WIDL")
                writeCom("WBTX2 013021150000020")
            ElseIf protocol.Equals("802.11g") Then
                writeCom("WIDL")
                writeCom("WBTX2 013041200000020")
            ElseIf protocol.Equals("802.11n") Then
                writeCom("WIDL")
                writeCom("WBTX2 013121200000020")
            End If
        End If
    End Sub
    Sub stopWifi(com As String)
        writeCom("WTFD 0")
        closeCom(com)
    End Sub
End Module
