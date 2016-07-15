Attribute VB_Name = "GPIB_CONTROLS"
Sub TEST_GPIB_MULTIMETER()
    Dim Mresult As Double
    Dim Mstatus As Boolean
    
'    MsgBox "Test DC Voltage measurement: GPIB::6"
'    Mresult = MM_DCV("GPIB::6", 10, "10", "0.001", 1)
'    MsgBox "Measured Average Voltage was: " & Mresult
'
'    MsgBox "Test DC Voltage measurement: GPIB::6"
'    Mresult = MM_DCV("GPIB::6", 10, "10", "0.001")
'    MsgBox "Measured Average Voltage was: " & Mresult
'
'    MsgBox "Test DC Voltage measurement: GPIB::6"
'    Mresult = MM_DCV("GPIB::6", 10, "10", "0.001", 2)
'    MsgBox "Measured Average Voltage was: " & Mresult
'
'    MsgBox "Test DC Voltage measurement: GPIB::6"
'    Mresult = MM_DCV("GPIB::6", 10, "10", "0.001", 3)
'    MsgBox "Measured Average Voltage was: " & Mresult
    
'
'    MsgBox "Test DC Resistance measurement: GPIB::3"
'    Mresult = MM_DCR("GPIB::3", 5)
'    MsgBox "Measured Average Resistance was: " & Mresult
    
    Mresult = MM_DCI("GPIB::3", 10, "0.1", "0.000001")
    MsgBox "Measured Average Current was: " & Mresult
End Sub

Sub TEST_GPIB_SUPPLY()
    Dim Result As Boolean
    
'    MsgBox "Initialize to 0V! GPIB-1"
'    result = Supply_Voltage("GPIB::1", 0, 0)
'    MsgBox "State is now " & result
'
'    MsgBox "Initialize 6V Supply to 0V! GPIB-2"
'    result = Supply_Voltage("GPIB::2", 0, 1)
'    MsgBox "State is now " & result
'
'    MsgBox "Initialize 25V Supply to 0V! GPIB-2"
'    result = Supply_Voltage("GPIB::2", 0, 2)
'    MsgBox "State is now " & result
'
'    MsgBox "2A Current Limit! GPIB-1"
'    result = Supply_Current("GPIB::1", 2, 0)
'    MsgBox "State is now " & result
'
'    MsgBox "1A Current Limit 6V Rail! GPIB-2"
'    result = Supply_Current("GPIB::2", 1, 1)
'    MsgBox "State is now " & result
'
'    MsgBox "1A Current Limit 25V Rail! GPIB-2"
'    result = Supply_Current("GPIB::2", 1, 2)
'    MsgBox "State is now " & result
'
'    MsgBox "Turn me on! GPIB::1"
'    result = Supply_On("GPIB::1")
'    MsgBox "State is now " & result
'
'    MsgBox "Turn me on! GPIB::2"
'    result = Supply_On("GPIB::2")
'    MsgBox "State is now " & result
'
'    MsgBox "3.6V! GPIB-1"
'    result = Supply_Voltage("GPIB::1", 3.6, 0)
'    MsgBox "State is now " & result
'
'    MsgBox "1.8V! 25V GPIB-2"
'    result = Supply_Voltage("GPIB::2", 1.8, 2)
'    MsgBox "State is now " & result
'
'    MsgBox "1.8V! 6V GPIB-2"
'    result = Supply_Voltage("GPIB::2", 1.8, 1)
'    MsgBox "State is now " & result
'
'    MsgBox "What a turn off! GPIB::1"
'    result = Supply_Off("GPIB::1")
'    MsgBox "State is now " & result
'
'    MsgBox "What a turn off! GPIB::2"
'    result = Supply_Off("GPIB::2")
'    MsgBox "State is now " & result
    
End Sub

Function MM_DCV(strGPIB_address, _
                AVERAGES As Integer, _
                Optional MRNG As String = "DEF", _
                Optional MRES As String = "DEF", _
                Optional MMOUT As Integer = 0) As Double

Dim ioMgr As VisaComLib.ResourceManager
Set ioMgr = New VisaComLib.ResourceManager

Dim instrument As VisaComLib.FormattedIO488
Set instrument = New VisaComLib.FormattedIO488
Set instrument.IO = ioMgr.Open(strGPIB_address)

Dim i As Integer
Dim CUR_ROUTE As String
Dim CUR_LENGTH As Integer
Dim CUR_NUMBER As Integer
Dim Msum As Double: Msum = 0
Dim MSTR As String: MSTR = ""

If MMOUT > 0 Then
    instrument.WriteString "ROUT:CLOS:STAT?"
    CUR_ROUTE = instrument.ReadString()
    CUR_LENGTH = Len(CUR_ROUTE)
    CUR_NUMBER = Int(Left(Right(CUR_ROUTE, CUR_LENGTH - 2), CUR_LENGTH - 4))
    If (CUR_NUMBER <> MMOUT) Then
        instrument.WriteString "rout:clos (@" & MMOUT & ")"
    End If
End If

If AVERAGES < 1 Then
    MM_DCV = 0
    GoTo end_MM_DCV
End If

For i = 1 To AVERAGES
    MSTR = "MEAS:VOLT:DC? " & MRNG & "," & MRES
    instrument.WriteString MSTR
    Msum = Msum + instrument.ReadString()
Next i
   
MM_DCV = Msum / AVERAGES
end_MM_DCV:

End Function

Function MM_DCR(strGPIB_address, _
                AVERAGES As Integer, _
                Optional MRNG As String = "DEF", _
                Optional MRES As String = "DEF") As Double

Dim ioMgr As VisaComLib.ResourceManager
Set ioMgr = New VisaComLib.ResourceManager

Dim instrument As VisaComLib.FormattedIO488
Set instrument = New VisaComLib.FormattedIO488
Set instrument.IO = ioMgr.Open(strGPIB_address)

Dim i As Integer
Dim Msum As Double: Msum = 0
Dim MSTR As String: MSTR = ""

If AVERAGES < 1 Then
    MM_DCR = 0
    GoTo end_MM_DCR
End If

For i = 1 To AVERAGES
    MSTR = "MEAS:RES? " & MRNG & "," & MRES
    instrument.WriteString MSTR
    Msum = Msum + instrument.ReadString()
Next i
   
MM_DCR = Msum / AVERAGES
end_MM_DCR:

End Function

Function MM_DCI(strGPIB_address, _
                AVERAGES As Integer, _
                Optional MRNG As String = "MIN", _
                Optional MRES As String = "DEF") As Double

Dim ioMgr As VisaComLib.ResourceManager
Set ioMgr = New VisaComLib.ResourceManager

Dim instrument As VisaComLib.FormattedIO488
Set instrument = New VisaComLib.FormattedIO488
Set instrument.IO = ioMgr.Open(strGPIB_address)

Dim i As Integer
Dim Msum As Double: Msum = 0
Dim MSTR As String: MSTR = ""

If AVERAGES < 1 Then
    MM_DCI = 0
    GoTo end_MM_DCI
End If

For i = 1 To AVERAGES
    MSTR = "MEAS:CURR:DC? " & MRNG & "," & MRES
    instrument.WriteString MSTR
    Msum = Msum + instrument.ReadString()
Next i
   
MM_DCI = Msum / AVERAGES
end_MM_DCI:

End Function

Function Supply_On(strGPIB_address)

Dim ioMgr As VisaComLib.ResourceManager
Set ioMgr = New VisaComLib.ResourceManager

Dim instrument As VisaComLib.FormattedIO488
Set instrument = New VisaComLib.FormattedIO488
Set instrument.IO = ioMgr.Open(strGPIB_address)

'instrument.WriteString "VOLTage " & str(dblSupply_Voltage)
instrument.WriteString "OUTP ON"
instrument.WriteString "OUTP:STATE?"
Supply_On = instrument.ReadString()

End Function

Function Supply_Off(strGPIB_address) ', dblSupply_Voltage)

Dim ioMgr As VisaComLib.ResourceManager
Set ioMgr = New VisaComLib.ResourceManager

Dim instrument As VisaComLib.FormattedIO488
Set instrument = New VisaComLib.FormattedIO488
Set instrument.IO = ioMgr.Open(strGPIB_address)

'instrument.WriteString "VOLTage " & str(dblSupply_Voltage)
instrument.WriteString "OUTP OFF"
instrument.WriteString "OUTP:STATE?"
Supply_Off = instrument.ReadString()

End Function

Function Supply_Voltage(strGPIB_address, SV As Double, Optional SUPOUT As Integer = 0) As Boolean

Dim ioMgr As VisaComLib.ResourceManager
Set ioMgr = New VisaComLib.ResourceManager

Dim instrument As VisaComLib.FormattedIO488
Set instrument = New VisaComLib.FormattedIO488
Set instrument.IO = ioMgr.Open(strGPIB_address)

If SUPOUT > 0 Then
    instrument.WriteString "INST:NSEL " & Str(SUPOUT)
End If

instrument.WriteString "VOLT " & Str(SV)
Supply_Voltage = True

End Function

Function Supply_Current(strGPIB_address, SC As Double, Optional SUPOUT As Integer = 0) As Boolean

Dim ioMgr As VisaComLib.ResourceManager
Set ioMgr = New VisaComLib.ResourceManager

Dim instrument As VisaComLib.FormattedIO488
Set instrument = New VisaComLib.FormattedIO488
Set instrument.IO = ioMgr.Open(strGPIB_address)

If SUPOUT > 0 Then
    instrument.WriteString "INST:NSEL " & Str(SUPOUT)
End If

instrument.WriteString "CURR " & Str(SC)
Supply_Current = True

End Function
