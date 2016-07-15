Attribute VB_Name = "DeltaDesign9015_"
Function DeltaDesign9015_Active(ByVal GPIB_Address As String)

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "active"

End Function

Function DeltaDesign9015_Standby(ByVal GPIB_Address As String)

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "standby"

End Function

Function DeltaDesign9015_SetTempSetpoint(ByVal GPIB_Address As String, ByVal TempSetpoint As Double)

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "setpoint " & CStr(TempSetpoint)

End Function

Function DeltaDesign9015_GetTempSetpoint(ByVal GPIB_Address As String, ByRef TempSetpoint As Double)

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "setpoint?"
    TempSetpoint = instrument.ReadString()

End Function

Function DeltaDesign9015_MeasureTemp(ByVal GPIB_Address As String, ByRef MeasureTemp As Double)

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "temperature?"
    TempSetpoint = instrument.ReadString()

End Function
