Attribute VB_Name = "Current_Load_Kikusui_"
Option Explicit
Option Base 0
Option Compare Text
'This Module will Control the KIKUSUI PLZ164W only.  This module requires that the Agilent IO Libraries Suite 16.3 be installed.
'   You can download and install the VISA from following web site:
'       http://www.home.agilent.com/en/pd-1985909-pn-E2094/io-libraries-suite-162?nid=-33330.977662.00&cc=US&lc=eng&cmpid=zzfindiosuite
'   After the Agilent IO Libraries Suite 16.3 is installed, it will install VISA that allows this module to function.  Go the "Tools" toolbar
'   and click on "References...".  Here you want to make sure "VISA COM 3.0 Type Library" is selected.
'
'       Below is a list of Functions that have been created for this module.  You can use this list and the search function to jump to a given
'       function.  Simply highlight or have your curser within the name of the function and hold [Ctrl] and [f] keys to bring up the search function.
'       you should see the name of the function you want to jump to is already in the search window.  Now just hit [Enter], and poof.  you are there.
'           List of Functions:
'                       Current_Load_Set_Current_Output
'                       Current_Load_Enable
'                       Current_Get_Load_Enable (new)
'                       Current_Load_Measure_Voltage
'                       Current_Load_Measure_Current
'                       Current_Load_Measure_Power
'                       Current_Load_Get_Current_Setting
'                       Current_Load_Set_Output_Mode
'
'
'       Modification Log: (Date, By, Modification)
'                           09-18-2013, Chris Sibley,   Original Version
'                           12-5-2013, Chris Sibley,    Added the Current_Get_Load_Enable function.
'
'


'********************************************************************************************************************************************************
' Function Current_Load_Set_Current_Output
'********************************************************************************************************************************************************
'   This function will set the current load to the output.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, Double, Current = This is the current limit level to be set to the output.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           09-18-2013, Chris Sibley,   Original Version
'
Function Current_Load_Set_Current_Output(ByVal GPIB_Address As String, ByVal Current As Double) As String
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "SOURce:CURRent:LEVel " & CStr(Current)

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 1) = "0" Then
        Current_Load_Set_Current_Output = "All Good"
    Else
        Current_Load_Set_Current_Output = Error_Check
    End If

End Function

'********************************************************************************************************************************************************
' Function Current_Load_Set_Voltage_Output
'********************************************************************************************************************************************************
'   This function will set the current load to the output.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, Double, Voltage = This is the voltage limit level to be set to the output.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           09-18-2013, Chris Sibley,   Original Version
'
Function Current_Load_Set_Voltage_Output(ByVal GPIB_Address As String, ByVal Voltage As Double) As String
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "SOURce:VOLTage:LEVel " & CStr(Voltage)

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        Current_Load_Set_Voltage_Output = "All Good"
    Else
        Current_Load_Set_Voltage_Output = Error_Check
    End If

End Function

'********************************************************************************************************************************************************
' Function Current_Load_Enable
'********************************************************************************************************************************************************
'   This function will enable/disable the outputs of the current load
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
'                           09-18-2013, Chris Sibley,   Original Version
'
Function Current_Load_Enable(ByVal GPIB_Address As String, ByVal On_Off As String) As String
    
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
        Current_Load_Enable = "All Good"
    Else
        Current_Load_Enable = Error_Check
    End If
    
End Function

'********************************************************************************************************************************************************
' Function Current_Get_Load_Enable
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
Function Current_Get_Load_Enable(ByVal GPIB_Address As String, ByRef On_Off As Boolean) As String
    
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
        Current_Get_Load_Enable = "All Good"
    Else
        Current_Get_Load_Enable = Error_Check
    End If
    
End Function
'********************************************************************************************************************************************************
' Function Current_Load_Measure_Voltage
'********************************************************************************************************************************************************
'   This function will request a voltage reading for a given output, and return the value through one of the arguments.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Measure_Voltage = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           09-18-2013, Chris Sibley,   Original Version
'
Function Current_Load_Measure_Voltage(ByVal GPIB_Address As String, ByRef Measure_Voltage As Double) As String

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
        Current_Load_Measure_Voltage = "All Good"
    Else
        Current_Load_Measure_Voltage = Error_Check
    End If
    
    
End Function

'********************************************************************************************************************************************************
' Function Current_Load_Measure_Current
'********************************************************************************************************************************************************
'   This function will request a Current reading for a given output, and return the value through one of the arguments.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Measure_Current = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           09-18-2013, Chris Sibley,   Original Version
'
Function Current_Load_Measure_Current(ByVal GPIB_Address As String, ByRef Measure_Current As Double) As String

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
        Current_Load_Measure_Current = "All Good"
    Else
        Current_Load_Measure_Current = Error_Check
    End If
    
    
End Function
'********************************************************************************************************************************************************
' Function Current_Load_Measure_Power
'********************************************************************************************************************************************************
'   This function will request a Current reading for a given output, and return the value through one of the arguments.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Measure_Power = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           09-18-2013, Chris Sibley,   Original Version
'
Function Current_Load_Measure_Power(ByVal GPIB_Address As String, ByRef Measure_Power As Double) As String

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "MEASure:POWer:DC? "
    Measure_Power = instrument.ReadString()
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Current_Load_Measure_Power = "All Good"
    Else
        Current_Load_Measure_Power = Error_Check
    End If
    
    
End Function
'********************************************************************************************************************************************************
' Function Current_Load_Get_Current_Setting
'********************************************************************************************************************************************************
'   This function will request the voltage setting, and return the value through one of the arguments.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Current_Setting = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           09-18-2013, Chris Sibley,   Original Version
'
Function Current_Load_Get_Current_Setting(ByVal GPIB_Address As String, ByRef Current_Setting As Double) As String

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim C_Setting As String
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "SOURce:CURRent:LEVel? "
    C_Setting = instrument.ReadString()
    
    Current_Setting = CDbl(Right(Left(C_Setting, 8), 7))
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Current_Load_Get_Current_Setting = "All Good"
    Else
        Current_Load_Get_Current_Setting = Error_Check
    End If
    
    
End Function
'********************************************************************************************************************************************************
' Function Current_Load_Set_Output_Mode
'********************************************************************************************************************************************************
'   This function will set the Mode of the output.  Since the KIKUSUI can not change modes while the output is on, this function will disable the output
'   before switching modes.  This function will keep the output off after finish running.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, Double, Voltage = This is the voltage level to be set to the given output.
'
'                   Optional, String, Mode = This is the Mode to be set to the output.  The Following are the only Modes this argument can except
'                                               "CC" for constant current
'                                               "CV" for constant voltage
'                                               "CP" for constant power
'                                               "CR" for constant resistance
'                                               "CCCV"  (reserved, not yet implimented)
'                                               "CRCV"  (reserved, not yet implimented)
'
'
'
'       Modification Log: (Date, By, Modification)
'                           09-18-2013, Chris Sibley,   Original Version
'
Function Current_Load_Set_Output_Mode(ByVal GPIB_Address As String, ByVal Mode As String) As String
    
    Dim Reply As String
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    Reply = Current_Load_Enable(GPIB_Address, "Off")
    
    instrument.WriteString "SOURce:FUNCtion:MODE " & Mode

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        Current_Load_Set_Output_Mode = "All Good"
    Else
        Current_Load_Set_Output_Mode = Error_Check
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
    Dim V As Double
'    Dim I As Integer
    
'    Set ioMgr = New VisaComLib.ResourceManager
    
'    Set instrument = New VisaComLib.FormattedIO488
'    Set instrument.IO = ioMgr.Open("GPIB::01")
'
'
'    instrument.WriteString "OUTPut:STATe?"  ' " & On_Off
'    Reply = instrument.ReadString()
    
    Reply = Current_Load_Set_Current_Output("GPIB::09", 0)
        
'    Reply = Current_Load_Enable("GPIB::09", "On")
    
'    Reply = Current_Load_Measure_Power("GPIB::09", V)
    
'    Reply = Current_Load_Set_Output_Mode("GPIB::09", "CV")

'    Reply = Current_Load_Set_Voltage_Output("GPIB::09", 4.5)
    
    'MsgBox CStr(V)
End Sub




'********************************************************************************************************************************************************
' Function Supply_Measure_Current
'********************************************************************************************************************************************************
'   This function will request a Current reading for a given output, and return the value through one of the arguments.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Output_Name = This sets the range of the measurnemt.  The valid srtrings can be:
'                                                   "P5V" = positive 6V output
'                                                   "P25V"  = positive 25V output
'                                                   "N25V" = negative 25V output
'                   Required, String, Measure_Current = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function Supply_Template(ByVal GPIB_Address As String, ByVal Output_Name As String, ByRef Measure_Current As Double) As String

    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "MEASure:CURRent:DC? " & Output_Name
    Measure_Current = instrument.ReadString()
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Supply_Template = "All Good"
    Else
        Supply_Template = Error_Check
    End If
    
    
End Function



