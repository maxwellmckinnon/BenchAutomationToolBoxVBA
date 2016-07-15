Attribute VB_Name = "Power_Supply_E3632A_"
Option Explicit
Option Base 0
Option Compare Text
'This Module will Control the Agilent E3631A only.  This module requires that the Agilent IO Libraries Suite 16.3 be installed.
'   You can download and install the VISA from following web site:
'       http://www.home.agilent.com/en/pd-1985909-pn-E2094/io-libraries-suite-162?nid=-33330.977662.00&cc=US&lc=eng&cmpid=zzfindiosuite
'   After the Agilent IO Libraries Suite 16.3 is installed, it will install VISA that allows this module to function.  Go the "Tools" toolbar
'   and click on "References...".  Here you want to make sure "VISA COM 3.0 Type Library" is selected.
'
'       Below is a list of Functions that have been created for this module.  You can use this list and the search function to jump to a given
'       function.  Simply highlight or have your curser within the name of the function and hold [Ctrl] and [f] keys to bring up the search function.
'       you should see the name of the function you want to jump to is already in the search window.  Now just hit [Enter], and poof.  you are there.
'           List of Functions:
'                       Supply_Set_Output
'                       Supply_Output_Enable
'                       Supply_Get_Output_Enable (new)
'                       Supply_Measure_Voltage
'                       Supply_Measure_Current
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'                           12-5-2013, Chris Sibley,    Added the Supply_Get_Output_Enable function.
'
'


'********************************************************************************************************************************************************
' Function Supply_Set_Output
'********************************************************************************************************************************************************
'   This function will set the voltage to the output.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'
'                   Required, Double, Voltage = This is the voltage level to be set to the given output.
'
'                   Optional, Double, Current = This is the current limit level to be set to the given output.  This is optional
'                                               the default is 1.0A
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-23-2013, Chris Sibley,   Created a copy module to control the A3632E supply.
'                           04-25-2013, Chris Sibley,   Original Version
'
Function Supply_Set_Output(ByVal GPIB_Address As String, ByVal Voltage As Double, Optional ByVal Current As Double = 1#) As String
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "VOLTage " & CStr(Voltage)

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        Supply_Set_Output = "All Good"
    Else
        Supply_Set_Output = Error_Check
    End If

End Function

'********************************************************************************************************************************************************
' Function Supply_Output_Enable
'********************************************************************************************************************************************************
'   This function will enable/disable the outputs of the power supply
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, On_Off = This will enable or disable the power supply output.
'                                                   "ON" = Turn the output on
'                                                   "OFF"  = Turn the output off
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-23-2013, Chris Sibley,   Created a copy module to control the A3632E supply.
'                           04-25-2013, Chris Sibley,   Original Version
'
Function Supply_Output_Enable(ByVal GPIB_Address As String, ByVal On_Off As String) As String
    
    Dim Reply As String
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "OUTPut:STATe?"
    Reply = instrument.ReadString()
    
    If On_Off = "On" Then
        If Left(Reply, 1) = "0" Then
            instrument.WriteString "OUTPut:STATe ON"
        End If
    Else
        If Left(Reply, 1) = "1" Then
            instrument.WriteString "OUTPut:STATe OFF"
        End If
    End If
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Supply_Output_Enable = "All Good"
    Else
        Supply_Output_Enable = Error_Check
    End If
    
End Function


'********************************************************************************************************************************************************
' Function Supply_Get_Output_Enable
'********************************************************************************************************************************************************
'   This function will check if the outputs of the power supply is enable or disable.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, On_Off = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'                                               true = enabled
'                                               false = disable
'
'
'
'       Modification Log: (Date, By, Modification)
'                           12-5-2013, Chris Sibley,   Original Version
'
Function Supply_Get_Output_Enable(ByVal GPIB_Address As String, ByRef On_Off As Boolean) As String
    
    Dim Reply As String
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "OUTPut:STATe?"
    Reply = instrument.ReadString()
    
    If Left(Reply, 1) = "0" Then
        On_Off = False
    Else
        On_Off = True
    End If
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Supply_Get_Output_Enable = "All Good"
    Else
        Supply_Get_Output_Enable = Error_Check
    End If
    
End Function

'********************************************************************************************************************************************************
' Function Supply_Measure_Voltage
'********************************************************************************************************************************************************
'   This function will request a voltage reading for for output, and return the value through one of the arguments.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'
'                   Required, String, Measure_Voltage = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-23-2013, Chris Sibley,   Created a copy module to control the A3632E supply.
'                           04-25-2013, Chris Sibley,   Original Version
'
Function Supply_Measure_Voltage(ByVal GPIB_Address As String, ByVal Voltage As Double, ByRef Measure_Voltage As Double) As String

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "MEASure:VOLTage:DC? "
    Measure_Voltage = instrument.ReadString()
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Supply_Measure_Voltage = "All Good"
    Else
        Supply_Measure_Voltage = Error_Check
    End If
    
    
End Function

'********************************************************************************************************************************************************
' Function Supply_Measure_Current
'********************************************************************************************************************************************************
'   This function will request a Current reading for the output, and return the value through one of the arguments.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'
'                   Required, String, Measure_Current = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-23-2013, Chris Sibley,   Created a copy module to control the A3632E supply.
'                           04-25-2013, Chris Sibley,   Original Version
'
Function Supply_Measure_Current(ByVal GPIB_Address As String, ByRef Measure_Current As Double) As String

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "MEASure:CURRent:DC? "
    Measure_Current = instrument.ReadString()
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Supply_Measure_Current = "All Good"
    Else
        Supply_Measure_Current = Error_Check
    End If
    
    
End Function

'********************************************************************************************************************************************************
' Function Supply_Get_Voltage_Setting
'********************************************************************************************************************************************************
'   This function will request the voltage setting, and return the value through one of the arguments.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'
'                   Required, String, Measure_Current = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-23-2013, Chris Sibley,   Original Version
'
Function Supply_Get_Voltage_Setting(ByVal GPIB_Address As String, ByRef Measure_Current As Double) As String

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "VOLTage? "
    Measure_Current = instrument.ReadString()
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Supply_Get_Voltage_Setting = "All Good"
    Else
        Supply_Get_Voltage_Setting = Error_Check
    End If
    
    
End Function



Private Function Error_Checker(ByVal GPIB_Address As String) As String
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim Reply As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Dim instrument As VisaComLib.FormattedIO488
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "SYSTem:ERRor?"
    Reply = instrument.ReadString()
    
    Error_Checker = Reply
End Function










'********************************************************************************************
'Test Program
Private Sub Tetest()
'    Dim ioMgr As VisaComLib.ResourceManager
'    Dim instrument As VisaComLib.FormattedIO488
'    Dim Error_Check As String
    Dim Reply As String
    Dim a As Double
    
'    Set ioMgr = New VisaComLib.ResourceManager
    
'    Set instrument = New VisaComLib.FormattedIO488
'    Set instrument.IO = ioMgr.Open("GPIB::01")
'
'
'    instrument.WriteString "OUTPut:STATe?"  ' " & On_Off
'    Reply = instrument.ReadString()
    
    Reply = Supply_Get_Voltage_Setting("GPIB::01", a)
'    Reply = Output_Enable("GPIB::01", "off")
        

End Sub


