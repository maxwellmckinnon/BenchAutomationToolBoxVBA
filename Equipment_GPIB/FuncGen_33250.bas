Attribute VB_Name = "FuncGen_33250"
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
'                       Func_Gen_Set_Output
'                       Func_Gen_Enable_Output
'
'
'
'
'
'
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2015, Chris Sibley,   Original Version
'
'


'********************************************************************************************************************************************************
' Function Func_Gen_Set_Output
'********************************************************************************************************************************************************
'   This function will set the voltage levels for the output signal.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'
'                   Required, Double, Voltage_High = This sets the higher level of the signal.
'
'                   Optional, Double, Voltage_Low = This sets the lower level of the signal
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2014, Chris Sibley,   Original Version
'
Function Func_Gen_Set_Output(ByVal GPIB_Address As String, ByVal Voltage_High As Double, Optional ByVal Voltage_Low As Double = 0) As String
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "VOLTage:HIGH " & CStr(Voltage_High)
    instrument.WriteString "VOLTage:LOW " & CStr(Voltage_Low)

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        Func_Gen_Set_Output = "All Good"
    Else
        Func_Gen_Set_Output = Error_Check
    End If

End Function


'********************************************************************************************************************************************************
' Function Func_Gen_Enable_Output
'********************************************************************************************************************************************************
'   This function will set the voltage to the given output.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Output_Name = This sets the range of the measurnemt.  The valid srtrings can be:
'                                                   "On" = Turn output on
'                                                   "Off"  = Turn output off
'
'
'
'       Modification Log: (Date, By, Modification)
'                           07-06-2015, Chris Sibley,   Original Version
'
Function Func_Gen_Enable_Output(ByVal GPIB_Address As String, ByVal State_str As String) As String
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open(GPIB_Address)

    instrument.WriteString "OUTput " & State_str
    

    Error_Check = Error_Checker(GPIB_Address)

    If Left(Error_Check, 2) = "+0" Then
        Func_Gen_Enable_Output = "All Good"
    Else
        Func_Gen_Enable_Output = Error_Check
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

Private Sub Test_Func_Gen_Set_Output()

    Call Func_Gen_Enable_Output("GPIB::16", "On")
    
End Sub

