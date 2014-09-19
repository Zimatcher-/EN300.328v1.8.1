Imports System.Xml

Module Sharp
    Sub loadCommands()
        Dim doc As XmlDocument = New XmlDocument()
        Dim node As XmlNode
        doc.Load(GlobalVar.sharpCommandsLocXML)
        node = doc.DocumentElement()
        ReDim GlobalVar.sharpCommands(0 To node.ChildNodes.Count - 1)
        For i = 0 To node.ChildNodes.Count - 1 Step 1
            sharpCommands(i) = node.ChildNodes.ItemOf(i).InnerText
        Next
    End Sub
    Function isSharpConnected()
        If isDeviceConnected("NB Series Serial Port") Then
            Return True
        Else
            Return False
        End If
    End Function
    Sub startWifi()
        loadCommands()
        openCom()
        writeCom(sharpCommands(0))
    End Sub
    Sub setChannel(freq As String, protocol As String)
        If protocol.Equals("802.11b") Then
            If freq.Equals("bottem") Then
                writeCom(sharpCommands(1))
            ElseIf freq.Equals("middle") Then
                writeCom(sharpCommands(2))
            ElseIf freq.Equals("top") Then
                writeCom(sharpCommands(3))
            End If
        ElseIf protocol.Equals("802.11g") Then
            If freq.Equals("bottem") Then
                writeCom(sharpCommands(4))
            ElseIf freq.Equals("middle") Then
                writeCom(sharpCommands(5))
            ElseIf freq.Equals("top") Then
                writeCom(sharpCommands(6))
            End If
        ElseIf protocol.Equals("802.11n") Then
            If freq.Equals("bottem") Then
                writeCom(sharpCommands(7))
            ElseIf freq.Equals("middle") Then
                writeCom(sharpCommands(8))
            ElseIf freq.Equals("top") Then
                writeCom(sharpCommands(9))
            End If
        End If
    End Sub
    Sub stopWifi()
        writeCom(sharpCommands(10))
        closeCom()
    End Sub
End Module
