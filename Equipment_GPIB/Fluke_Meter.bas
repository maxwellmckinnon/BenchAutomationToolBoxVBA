Attribute VB_Name = "Fluke_Meter"
Option Base 1
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Function ReadVoltage_Fluke(strGPIB_address)

    Dim ioMgr As VisaComLib.ResourceManager
    Set ioMgr = New VisaComLib.ResourceManager
    
    Dim instrument As VisaComLib.FormattedIO488
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(strGPIB_address)
    
    instrument.WriteString "MEAS:VOLT:DC? 10"
    v = Left(instrument.ReadString(), 7)
    ReadVoltage_Fluke = Cdouble(v)

End Function

Function ReadCurrent_Fluke(strGPIB_address)

    Dim ioMgr As VisaComLib.ResourceManager
    Set ioMgr = New VisaComLib.ResourceManager
    
    Dim instrument As VisaComLib.FormattedIO488
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(strGPIB_address)
    
    instrument.WriteString "MEAS:CURR:DC? 10"
    ReadCurrent_Fluke = Val(instrument.ReadString())

End Function

Function ReadAveVoltage_Fluke(strGPIB_address)

Dim ioMgr As VisaComLib.ResourceManager
    Set ioMgr = New VisaComLib.ResourceManager

    Dim instrument As VisaComLib.FormattedIO488
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(strGPIB_address)
    
    Dim n As Byte  'Number of samples
    Dim i As Byte
    Dim tempd1 As Double
    Dim tempd2 As Double
    Dim tempd3 As Double
    Dim tempd4 As Double
    Dim tempd5 As Double
    
    instrument.WriteString "MEAS:VOLT:DC? 10"
    tempd1 = CDbl(instrument.ReadString())
    instrument.WriteString "MEAS:VOLT:DC? 10"
    tempd2 = CDbl(instrument.ReadString())
    instrument.WriteString "MEAS:VOLT:DC? 10"
    tempd3 = CDbl(instrument.ReadString())
    instrument.WriteString "MEAS:VOLT:DC? 10"
    tempd4 = CDbl(instrument.ReadString())
    instrument.WriteString "MEAS:VOLT:DC? 10"
    tempd5 = CDbl(instrument.ReadString())
    
    ReadAveVoltage_Fluke = Round((tempd1 + tempd2 + tempd3 + tempd4 + tempd5) / 5, 6)
                
End Function

Function ReadAveNVoltage_Fluke(strGPIB_address)

Dim ioMgr As VisaComLib.ResourceManager
    Set ioMgr = New VisaComLib.ResourceManager

    Dim instrument As VisaComLib.FormattedIO488
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(strGPIB_address)
    
    Dim n As Double  'Number of samples
    Dim i As Double
    Dim tempd As Double
    tempd = 0
    n = 3
        
    For i = 1 To n
        instrument.WriteString "MEAS:VOLT:DC?"
        tempd = tempd + CDbl(instrument.ReadString())
    Next i
    
    ReadAveNVoltage_Fluke = tempd / n
                
End Function

Function ReadAve_Fluke(strGPIB_address)


Dim ioMgr As VisaComLib.ResourceManager
    Set ioMgr = New VisaComLib.ResourceManager

    Dim instrument As VisaComLib.FormattedIO488
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(strGPIB_address)
    
    Dim n As Double  'Number of samples
    Dim i As Double
    Dim tempd As Double
    tempd = 0
    n = 5
        
    For i = 1 To n
        Sleep (50)
        instrument.WriteString "READ?"
        Sleep (50)
        tempd = tempd + CDbl(instrument.ReadString())
    Next i
    
    ReadAve_Fluke = tempd / n
                
End Function


