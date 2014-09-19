Module GlobalVar
    Public location As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\Desktop\"
    Public freq As String() = {"", "", ""}
    Public tests As String() = {"Power Spectral Density", "Occupied Bandwidth", "Out Band Emissions"}
    Public protocols As String() = {"802.11b", "802.11g", "802.11n"}
    Public temps As String() = {"-10", "Ambient", "+55"}
    Public jobNo As String = ""
    Public offset As String = "31.2dB"
    Public psd As String() = {"2400MHz", "2483.5MHz", "10dBm", "5 DB", "10 KHz", "30 KHz", "10S", "30001", "RMS", "MAXH"}
    'PSD TEST FIELDS = {start, stop, ref, attenuation, offset, ResBandWidth, VidBandWidth, SweepTime, SweepPoints, Detector, Trace}
    Public obe As String() = {"2399.5", "2484", "0Hz", "10dBm", "5 DB", "1 MHz", "3 MHz", "20ms", "5000", "100", "RMS", "MAXH", "CFIL"}
    'OBE TEST FIELDS = {}{bottemBand, topBand, Span, ref, attenuation, ResBandWidth, VidBandWidth, SweepTime, SweepPoints, SweepCount, Detector, Trace, Filter}
    Public obw As String() = {"2412MHz", "2472MHz", "40Mhz", "20dBm", "5 DB", "500 KHz", "2 MHz", "10ms", "625", "100", "RMS", "MAXH"}
    'PSD TEST FIELDS = {bottemCentre, topCentre, Span, ref, attenuation, ResBandWidth, VidBandWidth, SweepTime, SweepPoints, SweepCount, Detector, Trace}
    Public commands As String()
    Public commandsLocXML As String = "spectrumAnalyzerCommands.XML"
    Public sharpConnected As Boolean = False
    Public sharpCommands As String()
    Public sharpCommandsLocXML As String = "sharpCommands.XML"

End Module
