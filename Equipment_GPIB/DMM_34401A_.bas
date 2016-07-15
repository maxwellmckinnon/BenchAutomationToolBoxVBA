Attribute VB_Name = "DMM_34401A_"
Option Explicit
Option Base 0
Option Compare Text
'This Module will Control the Agilent 34401A only.  This module requires that the Agilent IO Libraries Suite 16.3 be installed.
'   You can download and install the VISA from following web site:
'       http://www.home.agilent.com/en/pd-1985909-pn-E2094/io-libraries-suite-162?nid=-33330.977662.00&cc=US&lc=eng&cmpid=zzfindiosuite
'   After the Agilent IO Libraries Suite 16.3 is installed, it will install VISA that allows this module to function.  Go the "Tools" toolbar
'   and click on "References...".  Here you want to make sure "VISA COM 3.0 Type Library" is selected.
'
'       Below is a list of Functions that have been created for this module.  You can use this list and the search function to jump to a given
'       function.  Simply highlight or have your curser within the name of the function and hold [Ctrl] and [f] keys to bring up the search function.
'       you should see the name of the function you want to jump to is already in the search window.  Now just hit [Enter], and poof.  you are there.
'           List of Functions:
'                       DMM_Config_DC_Volt
'                       DMM_Config_AC_Volt
'                       DMM_Config_DC_Current
'                       DMM_Config_AC_Current
'                       DMM_Get_Config
'                       DMM_Set_Distplay
'                       DMM_Get_Null
'                       DMM_Get_Reading
'                       DMM_Set_Band
'                       DMM_Get_Band
'                       DMM_Clear_Reset
'                       DMM_Set_Impedance
'                       DMM_Write_Display
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
'


'********************************************************************************************************************************************************
' Function DMM_Config_DC_Volt
'********************************************************************************************************************************************************
'   This function will set the DMM to measure DC Voltage. The Range and Resolution can be set through this function, but they're optional.
'       The Range and Resolution is defaulted to auto.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Optional, String, Range_ = This sets the range of the measurnemt.  The valid srtrings can be:
'                                              "DEF" (Default) - auto range
'                                              "MIN" or "MINimum" = 0.10
'                                              "MAX" or "MAXimum" = 1.0e3
'                                              Cstr(Double) - from 1e-6 to 10
'                   Optional, String, Resolution = This sets the range of the measurnemt.  The valid srtrings can be:
'                                                  "DEF" (Default) - auto
'                                                  "MIN" or "MINimum" = 3.0e-8
'                                                  "MAX" or "MAXimum" = 9.99999e-2
'                                                  Cstr(Double) - from 3.0e-8 to 9.99999e-2
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Config_DC_Volt(ByVal GPIB_Address As String, Optional ByVal Range_ As String = "DEF", Optional ByVal Resolution As String = "DEF") As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "CONFigure:VOLTage:DC " & Range_ & ", " & Resolution

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        DMM_Config_DC_Volt = "All Good"
    Else
        DMM_Config_DC_Volt = Error_Check
    End If

End Function

'********************************************************************************************************************************************************
' Function DMM_Config_AC_Volt
'********************************************************************************************************************************************************
'   This function will set the DMM to measure AC Voltage. The Range and Resolution can be set through this function, but they're optional.
'       The Range and Resolution is defaulted to auto.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Optional, String, Range_ = This sets the range of the measurnemt.  The valid srtrings can be:
'                                              "DEF" (Default) - auto range
'                                              "MIN" or "MINimum" = 1.00e-1
'                                              "MAX" or "MAXimum" = 1.00e+2
'                                              Cstr(Double) - from 1e-1 to 100
'                   Optional, String, Resolution = This sets the range of the measurnemt.  The valid srtrings can be:
'                                                  "DEF" (Default) - auto
'                                                  "MIN" or "MINimum" = 1.00e-7
'                                                  "MAX" or "MAXimum" = 9.99999e-2
'                                                  Cstr(Double) - from 1.0e-7 to 9.99999e-2
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Config_AC_Volt(ByVal GPIB_Address As String, Optional ByVal Range_ As String = "DEF", Optional ByVal Resolution As String = "DEF") As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "CONFigure:VOLTage:AC " & Range_ & ", " & Resolution

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        DMM_Config_AC_Volt = "All Good"
    Else
        DMM_Config_AC_Volt = Error_Check
    End If

End Function

'********************************************************************************************************************************************************
' Function DMM_Config_DC_Current
'********************************************************************************************************************************************************
'   This function will set the DMM to measure DC Current. The Range and Resolution can be set through this function, but they're optional.
'       The Range and Resolution is defaulted to auto.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Optional, String, Range_ = This sets the range of the measurnemt.  The valid srtrings can be:
'                                              "DEF" (Default) - auto range
'                                              "MIN" or "MINimum" = 1.00e-2
'                                              "MAX" or "MAXimum" = 3.00e0
'                                              Cstr(Double) - from 1.00e-2 to 3.00e0
'                   Optional, String, Resolution = This sets the range of the measurnemt.  The valid srtrings can be:
'                                                  "DEF" (Default) - auto
'                                                  "MIN" or "MINimum" = 3.00e-9
'                                                  "MAX" or "MAXimum" = 3.00e-4
'                                                  Cstr(Double) - from 3.00e-9 to 3.00e-4
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Config_DC_Current(ByVal GPIB_Address As String, Optional ByVal Range_ As String = "DEF", Optional ByVal Resolution As String = "DEF") As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "CONFigure:CURRent:DC " & Range_ & ", " & Resolution

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        DMM_Config_DC_Current = "All Good"
    Else
        DMM_Config_DC_Current = Error_Check
    End If

End Function

'********************************************************************************************************************************************************
' Function DMM_Config_AC_Current
'********************************************************************************************************************************************************
'   This function will set the DMM to measure AC Current. The Range and Resolution can be set through this function, but they're optional.
'       The Range and Resolution is defaulted to auto.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Optional, String, Range_ = This sets the range of the measurnemt.  The valid srtrings can be:
'                                              "DEF" (Default) - auto range
'                                              "MIN" or "MINimum" = 1.00
'                                              "MAX" or "MAXimum" = 3.00
'                                              Cstr(Double) - from 1.00 to 3.00
'                   Optional, String, Resolution = This sets the range of the measurnemt.  The valid srtrings can be:
'                                                  "DEF" (Default) - auto
'                                                  "MIN" or "MINimum" = 1.00e-6
'                                                  "MAX" or "MAXimum" = 3.00e-4
'                                                  Cstr(Double) - from 1.0e-6 to 3.00e-4
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Config_AC_Current(ByVal GPIB_Address As String, Optional ByVal Range_ As String = "DEF", Optional ByVal Resolution As String = "DEF") As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "CONFigure:CURRent:AC " & Range_ & ", " & Resolution

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        DMM_Config_AC_Current = "All Good"
    Else
        DMM_Config_AC_Current = Error_Check
    End If

End Function

'********************************************************************************************************************************************************
' Function DMM_Get_Config
'********************************************************************************************************************************************************
'   This function will get the DMM configuration.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Configuration = This is sent by reference to return the string that is read from the equipment.
'                                                     This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Get_Config(ByVal GPIB_Address As String, ByRef Configuration As String) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "CONFigure?"
    Configuration = instrument.ReadString()
    
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        DMM_Get_Config = "All Good"
    Else
        DMM_Get_Config = Error_Check
    End If
    
End Function


'********************************************************************************************************************************************************
' Function DMM_Set_Distplay
'********************************************************************************************************************************************************
'   This function will set the display on the DMM to any string of 12 Characters or less.  The string will display on the equipment and then a
'       message box will appear that shows the routine is paused.  As soon the user clicks ok, the display will clear and return to normal.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                                    The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Message = This a string of charatcers to be displayed on the unit.  It will only dispaly the first 12
'                                               characters.  The upper case and lower case is limited on the type of display the DMM has.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Set_Distplay(ByVal GPIB_Address As String, ByRef Message As String) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "DISPlay:TEXT '" & Message & "'"
    
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        DMM_Set_Distplay = "All Good"
    Else
        DMM_Set_Distplay = Error_Check
    End If
    
    MsgBox "Pause", vbOKOnly, "Display On"
    
End Function

'********************************************************************************************************************************************************
' Function DMM_Get_Null
'********************************************************************************************************************************************************
'   This function will get the null offset of the DMM.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                                    The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Null_Offset = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Get_Null(ByVal GPIB_Address As String, ByRef Null_Offset As String) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "CALCulate:FUNCtion NULL"
    instrument.WriteString "CALCulate:NULL:OFFSet MAX"
    instrument.WriteString "CALCulate:NULL:OFFSet?"
    Null_Offset = instrument.ReadString()
    
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        DMM_Get_Null = "All Good"
    Else
        DMM_Get_Null = Error_Check
    End If
    
End Function

'********************************************************************************************************************************************************
' Function DMM_Get_Reading
'********************************************************************************************************************************************************
'   This function will get the measured reading of the DMM.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                                    The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Reading = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'                   Optional, Integer, Number_of_Averages_int = this sets the number of readings that will be averaged for the reading that will be returned.
'                                                   default = 1
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'                           12-21-2015, Chris Sibley,   adding the optional argument for averaging.
'
Function DMM_Get_Reading(ByVal GPIB_Address As String, ByRef Reading As Double, Optional ByVal Number_of_Averages_int As Integer = 1) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Dim Str_Reading As String
    Dim Reading_Array_dbl() As Double
    Dim Running_Reading_dbl As Double
    Dim i As Integer
    ReDim Reading_Array_dbl(Number_of_Averages_int) As Double
    
    Running_Reading_dbl = 0
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    For i = 1 To Number_of_Averages_int Step 1
        instrument.WriteString "READ?"
        Str_Reading = instrument.ReadString()
        Reading_Array_dbl(i) = CDbl(Str_Reading)
        
        Error_Check = Error_Checker(GPIB_Address)
        
        If Left(Error_Check, 2) = "+0" Then
            DMM_Get_Reading = "All Good"
        Else
            DMM_Get_Reading = Error_Check
            Exit Function
        End If
    Next i

    For i = 1 To Number_of_Averages_int Step 1
        Running_Reading_dbl = Running_Reading_dbl + Reading_Array_dbl(i)
    Next i
    Reading = Running_Reading_dbl / CDbl(Number_of_Averages_int)
End Function

'********************************************************************************************************************************************************
' Function DMM_Set_Band
'********************************************************************************************************************************************************
'   This function will set the bandwidth for the measurement of the DMM.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                  The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, BAND = sets the bandwidth of the measurement.  The three options are:
'                          "3" Hz, it takes longest amount of time per measurements
'                          "20" Hz, the default
'                          "200" Hz, it takes shortest amount of time per measurements
'                          "MIN" or "MINimum" = 3Hz
'                          "MAX" or MAXimum" = 200Hz
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Set_Band(ByVal GPIB_Address As String, ByVal Band As String) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
   
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "DETector:BANDwidth " & Band
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        DMM_Set_Band = "All Good"
    Else
        DMM_Set_Band = Error_Check
    End If
    
End Function

'********************************************************************************************************************************************************
' Function DMM_Get_Band
'********************************************************************************************************************************************************
'   This function will get the bandwidth for the measurement of the DMM.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                  The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, BAND = This is sent by reference to return the string that is read from the equipment.
'                                            This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Get_Band(ByVal GPIB_Address As String, ByRef Band As Double) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Dim Str_Reading As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "DETector:BANDwidth?"
    Str_Reading = instrument.ReadString()
    Band = CDbl(Str_Reading)
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        DMM_Get_Band = "All Good"
    Else
        DMM_Get_Band = Error_Check
    End If
    
End Function

'********************************************************************************************************************************************************
' Function DMM_Clear_Reset
'********************************************************************************************************************************************************
'   This function will clear and reset the DMM.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                  The address is ranged between "GPIB::00" to "GPIB::31"
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Clear_Reset(ByVal GPIB_Address As String) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Dim Str_Reading As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "*RST"
    instrument.WriteString "*CLS"
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        DMM_Clear_Reset = "All Good"
    Else
        DMM_Clear_Reset = Error_Check
    End If
    
End Function

'********************************************************************************************************************************************************
' Private Function Error_Checker
'********************************************************************************************************************************************************
'   This function will check to see if there are any error reported on the DMM.  It will return the error results.  this function can not be called
'   outside this module.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                  The address is ranged between "GPIB::00" to "GPIB::31"
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Private Function Error_Checker(ByVal GPIB_Address As String) As String
    '
    '
    '   Defining the local Variables
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


'********************************************************************************************************************************************************
' Function DMM_Set_Impedance
'********************************************************************************************************************************************************
'   This function sets the input impedance of the DMM.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                  The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, On_Off = this will either enable or disable:
'                                              "ON" = set to high impeadance
'                                              "OFF" = Default impeadance
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function DMM_Set_Impedance(ByVal GPIB_Address As String, ByVal On_Off As String) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "INPut:IMPedance:AUTO " & On_Off      ' inser command between
    
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        DMM_Set_Impedance = "All Good"               '
    Else
        DMM_Set_Impedance = Error_Check
    End If
    
End Function

'********************************************************************************************************************************************************
' Function DMM_Write_Display
'********************************************************************************************************************************************************
'   This Function will clear the text on the display
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                  The address is ranged between "GPIB::00" to "GPIB::31"
'                   Optional, String, Text_str = This sets the display to display the string that is passed, or clears the display to allow the
'                                   display to show the measurement by passing the string "Clear" (this is the defualt)
'                   Optional, String, State_str = This turns on or off the display.  The defualt is "on"
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-14-2015, Chris Sibley,   Original Version
'
Function DMM_Write_Display(ByVal GPIB_Address As String, Optional ByVal Text_str As String = "Clear", Optional ByVal State_str As String = "On") As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String

    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    If Text_str = "Clear" Then
        instrument.WriteString "DISPlay:TEXT:CLEAR"
    Else
        instrument.WriteString "DISPlay:TEXT " & Chr(34) & Text_str & Chr(34)
    End If
    
    If State_str = "On" Then
        instrument.WriteString "DISPlay On"
    Else
        instrument.WriteString "DISPlay Off"
    End If
    
    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        DMM_Write_Display = "All Good"               '
    Else
        DMM_Write_Display = Error_Check
    End If

End Function




'********************************************************************************************************************************************************
' Function Template
'********************************************************************************************************************************************************
'   This function is a template for developers to add more functionality to this module.  Add the oppropiate arguments that needs to be passed.
'       If you're reading from the instrument, make sure you pass an argument as ByRef so it can return the string that was request from the DMM.
'       To pass an argument that is optional, it has to be delared as "Optional (ByVal|ByRef) {name} as {Data Type} = {default value}".
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                  The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Reading = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Function Template(ByVal GPIB_Address As String, ByRef Reading As String) As String
    '
    '
    '   Defining the local Variables
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    
    Set ioMgr = New VisaComLib.ResourceManager
    
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)
    
    instrument.WriteString "*IDN?"       ' inser command between
    Reading = instrument.ReadString()
    
    
    Error_Check = Error_Checker(GPIB_Address)
    
    If Left(Error_Check, 2) = "+0" Then
        Template = "All Good"               '
    Else
        Template = Error_Check
    End If
    
End Function




'INPut:IMPedance:AUTO {OFF|ON}



'********************************************************************************************
'Test Program
Private Sub Tetest()
'
    Dim Reply As String
    Dim Check As String
    Dim a As Double
    
    Reply = Template("GPIB::00", Check)
End Sub

Private Sub Test_DMM_Write_Display()

    Call DMM_Write_Display("GPIB::11")
End Sub

